import OAuthProvider from "@cloudflare/workers-oauth-provider";
import { McpAgent } from "agents/mcp";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";

// ---------------------------------------------------------------------------
// Config
// ---------------------------------------------------------------------------
// "consumers" = personal Microsoft accounts only (Outlook.com, Hotmail, Live).
const MS_AUTH = "https://login.microsoftonline.com/consumers/oauth2/v2.0";
const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = "offline_access User.Read Mail.Read Files.Read Tasks.ReadWrite";

interface Env {
	OAUTH_KV: KVNamespace;
	OAUTH_PROVIDER: any; // injected by OAuthProvider
	MCP_OBJECT: DurableObjectNamespace;
	MICROSOFT_CLIENT_ID: string;
	MICROSOFT_CLIENT_SECRET: string;
	ALLOWED_EMAIL?: string; // optional: lock the server to your own account
}

type Props = {
	userId: string;
	email: string;
	refreshToken: string;
	accessToken: string;
	expiresAt: number; // epoch ms
};

// ---------------------------------------------------------------------------
// MCP server
// ---------------------------------------------------------------------------
export class MicrosoftMCP extends McpAgent<Env, unknown, Props> {
	server = new McpServer({ name: "Personal Microsoft", version: "1.0.0" });

	private cached?: { token: string; expiresAt: number };

	private async token(): Promise<string> {
		const now = Date.now();
		const c = this.cached ?? {
			token: this.props!.accessToken,
			expiresAt: this.props!.expiresAt,
		};
		if (c.expiresAt - 60_000 > now) return c.token;

		const res = await fetch(`${MS_AUTH}/token`, {
			method: "POST",
			headers: { "Content-Type": "application/x-www-form-urlencoded" },
			body: new URLSearchParams({
				client_id: this.env.MICROSOFT_CLIENT_ID,
				client_secret: this.env.MICROSOFT_CLIENT_SECRET,
				grant_type: "refresh_token",
				refresh_token: this.props!.refreshToken,
				scope: SCOPES,
			}),
		});
		if (!res.ok) {
			throw new Error(
				`Microsoft token refresh failed (${res.status}). Reconnect the connector.`,
			);
		}
		const t = (await res.json()) as { access_token: string; expires_in: number };
		this.cached = {
			token: t.access_token,
			expiresAt: Date.now() + t.expires_in * 1000,
		};
		return t.access_token;
	}

	private async graph(path: string, init: RequestInit = {}): Promise<Response> {
		const res = await fetch(path.startsWith("http") ? path : `${GRAPH}${path}`, {
			...init,
			headers: {
				Authorization: `Bearer ${await this.token()}`,
				...(init.headers ?? {}),
			},
		});
		if (!res.ok) {
			const body = await res.text();
			throw new Error(`Graph ${res.status}: ${body.slice(0, 500)}`);
		}
		return res;
	}

	private async graphJson(path: string, init: RequestInit = {}) {
		return (await this.graph(path, init)).json() as Promise<any>;
	}

	private ok(data: unknown) {
		return { content: [{ type: "text" as const, text: JSON.stringify(data, null, 2) }] };
	}

	private fail(e: unknown) {
		return {
			isError: true,
			content: [{ type: "text" as const, text: e instanceof Error ? e.message : String(e) }],
		};
	}

	async init() {
		// ----------------------------- Outlook ------------------------------
		this.server.tool(
			"outlook_email_search",
			"Search or list emails in the personal Outlook mailbox. Omit query to list the newest messages.",
			{
				query: z.string().optional().describe("Search text (sender, subject, body)"),
				folder: z
					.enum(["inbox", "sentitems", "drafts", "archive", "deleteditems", "junkemail"])
					.optional()
					.describe("Limit to one folder; default is all mail"),
				top: z.number().int().min(1).max(25).default(10),
			},
			async ({ query, folder, top }) => {
				try {
					const base = folder ? `/me/mailFolders/${folder}/messages` : "/me/messages";
					const select =
						"id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments";
					const qs = query
						? `?$search=${encodeURIComponent(`"${query.replace(/"/g, "")}"`)}&$top=${top}&$select=${select}`
						: `?$orderby=receivedDateTime desc&$top=${top}&$select=${select}`;
					const data = await this.graphJson(base + qs);
					return this.ok(
						data.value.map((m: any) => ({
							id: m.id,
							subject: m.subject,
							from: m.from?.emailAddress,
							received: m.receivedDateTime,
							unread: !m.isRead,
							hasAttachments: m.hasAttachments,
							preview: m.bodyPreview,
						})),
					);
				} catch (e) {
					return this.fail(e);
				}
			},
		);

		this.server.tool(
			"outlook_email_read",
			"Read the full text of one email by id (from outlook_email_search).",
			{ id: z.string() },
			async ({ id }) => {
				try {
					const m = await this.graphJson(
						`/me/messages/${encodeURIComponent(id)}?$select=subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments`,
						{ headers: { Prefer: 'outlook.body-content-type="text"' } },
					);
					return this.ok({
						subject: m.subject,
						from: m.from?.emailAddress,
						to: m.toRecipients?.map((r: any) => r.emailAddress),
						cc: m.ccRecipients?.map((r: any) => r.emailAddress),
						received: m.receivedDateTime,
						hasAttachments: m.hasAttachments,
						body: (m.body?.content ?? "").slice(0, 20_000),
					});
				} catch (e) {
					return this.fail(e);
				}
			},
		);

		// ----------------------------- OneDrive -----------------------------
		this.server.tool(
			"onedrive_search",
			"Search OneDrive files and folders by name or content.",
			{ query: z.string(), top: z.number().int().min(1).max(25).default(10) },
			async ({ query, top }) => {
				try {
					const q = encodeURIComponent(query.replace(/'/g, "''"));
					const data = await this.graphJson(
						`/me/drive/root/search(q='${q}')?$top=${top}&$select=id,name,size,lastModifiedDateTime,webUrl,file,folder,parentReference`,
					);
					return this.ok(data.value.map(summarizeItem));
				} catch (e) {
					return this.fail(e);
				}
			},
		);

		this.server.tool(
			"onedrive_list_folder",
			"List items in a OneDrive folder. Omit path for the root, e.g. 'Documents/Notes'.",
			{ path: z.string().optional(), top: z.number().int().min(1).max(100).default(50) },
			async ({ path, top }) => {
				try {
					const clean = path?.replace(/^\/+|\/+$/g, "");
					const base = clean
						? `/me/drive/root:/${clean.split("/").map(encodeURIComponent).join("/")}:/children`
						: "/me/drive/root/children";
					const data = await this.graphJson(
						`${base}?$top=${top}&$select=id,name,size,lastModifiedDateTime,webUrl,file,folder`,
					);
					return this.ok(data.value.map(summarizeItem));
				} catch (e) {
					return this.fail(e);
				}
			},
		);

		this.server.tool(
			"onedrive_read_file",
			"Read a text-based OneDrive file (txt, md, csv, json, code, etc.) by id. Office docs and PDFs are not parsed.",
			{ id: z.string(), max_chars: z.number().int().min(500).max(50_000).default(20_000) },
			async ({ id, max_chars }) => {
				try {
					const meta = await this.graphJson(
						`/me/drive/items/${encodeURIComponent(id)}?$select=name,size,file`,
					);
					const mime: string = meta.file?.mimeType ?? "";
					const textual =
						mime.startsWith("text/") ||
						/json|xml|csv|javascript|yaml|markdown/.test(mime) ||
						/\.(txt|md|csv|json|ya?ml|xml|js|ts|py|html|css|log|ini|toml)$/i.test(meta.name);
					if (!textual) {
						return this.ok({
							name: meta.name,
							note: `Not readable as plain text (${mime || "unknown type"}).`,
						});
					}
					if (meta.size > 5_000_000) {
						return this.ok({ name: meta.name, note: "File is larger than 5 MB." });
					}
					// Graph answers with a 302 to a pre-authenticated URL; fetch follows it.
					const res = await this.graph(`/me/drive/items/${encodeURIComponent(id)}/content`);
					const text = await res.text();
					return this.ok({
						name: meta.name,
						truncated: text.length > max_chars,
						content: text.slice(0, max_chars),
					});
				} catch (e) {
					return this.fail(e);
				}
			},
		);

		// ------------------------------ To Do -------------------------------
		this.server.tool("todo_list_lists", "List your Microsoft To Do lists.", {}, async () => {
			try {
				const data = await this.graphJson("/me/todo/lists");
				return this.ok(
					data.value.map((l: any) => ({ id: l.id, name: l.displayName, wellKnown: l.wellknownListName })),
				);
			} catch (e) {
				return this.fail(e);
			}
		});

		this.server.tool(
			"todo_list_tasks",
			"List tasks in a To Do list (open tasks by default).",
			{
				list_id: z.string(),
				include_completed: z.boolean().default(false),
				top: z.number().int().min(1).max(100).default(50),
			},
			async ({ list_id, include_completed, top }) => {
				try {
					const filter = include_completed
						? ""
						: `&$filter=${encodeURIComponent("status ne 'completed'")}`;
					const data = await this.graphJson(
						`/me/todo/lists/${encodeURIComponent(list_id)}/tasks?$top=${top}${filter}`,
					);
					return this.ok(
						data.value.map((t: any) => ({
							id: t.id,
							title: t.title,
							status: t.status,
							importance: t.importance,
							due: t.dueDateTime?.dateTime,
							notes: t.body?.content || undefined,
						})),
					);
				} catch (e) {
					return this.fail(e);
				}
			},
		);

		this.server.tool(
			"todo_create_task",
			"Create a task in a To Do list.",
			{
				list_id: z.string(),
				title: z.string(),
				due_date: z.string().regex(/^\d{4}-\d{2}-\d{2}$/).optional().describe("YYYY-MM-DD"),
				notes: z.string().optional(),
				importance: z.enum(["low", "normal", "high"]).default("normal"),
			},
			async ({ list_id, title, due_date, notes, importance }) => {
				try {
					const body: any = { title, importance };
					if (due_date) body.dueDateTime = { dateTime: `${due_date}T00:00:00`, timeZone: "UTC" };
					if (notes) body.body = { content: notes, contentType: "text" };
					const t = await this.graphJson(
						`/me/todo/lists/${encodeURIComponent(list_id)}/tasks`,
						{
							method: "POST",
							headers: { "Content-Type": "application/json" },
							body: JSON.stringify(body),
						},
					);
					return this.ok({ id: t.id, title: t.title, status: t.status });
				} catch (e) {
					return this.fail(e);
				}
			},
		);

		this.server.tool(
			"todo_complete_task",
			"Mark a To Do task as completed.",
			{ list_id: z.string(), task_id: z.string() },
			async ({ list_id, task_id }) => {
				try {
					const t = await this.graphJson(
						`/me/todo/lists/${encodeURIComponent(list_id)}/tasks/${encodeURIComponent(task_id)}`,
						{
							method: "PATCH",
							headers: { "Content-Type": "application/json" },
							body: JSON.stringify({ status: "completed" }),
						},
					);
					return this.ok({ id: t.id, title: t.title, status: t.status });
				} catch (e) {
					return this.fail(e);
				}
			},
		);
	}
}

function summarizeItem(i: any) {
	return {
		id: i.id,
		name: i.name,
		type: i.folder ? "folder" : "file",
		size: i.size,
		modified: i.lastModifiedDateTime,
		mimeType: i.file?.mimeType,
		path: i.parentReference?.path?.replace("/drive/root:", "") || undefined,
		url: i.webUrl,
	};
}

// ---------------------------------------------------------------------------
// Microsoft sign-in (the "default handler" of the OAuth provider)
// ---------------------------------------------------------------------------
const COOKIE = "ms_oauth_state";

const authHandler = {
	async fetch(request: Request, env: Env): Promise<Response> {
		const url = new URL(request.url);

		if (url.pathname === "/authorize") {
			// Claude -> this server. Remember its request, then bounce to Microsoft.
			const oauthReqInfo = await env.OAUTH_PROVIDER.parseAuthRequest(request);
			if (!oauthReqInfo.clientId) return new Response("Invalid request", { status: 400 });

			const state = crypto.randomUUID();
			await env.OAUTH_KV.put(`ms_state:${state}`, JSON.stringify(oauthReqInfo), {
				expirationTtl: 600,
			});

			const ms = new URL(`${MS_AUTH}/authorize`);
			ms.searchParams.set("client_id", env.MICROSOFT_CLIENT_ID);
			ms.searchParams.set("response_type", "code");
			ms.searchParams.set("redirect_uri", `${url.origin}/callback`);
			ms.searchParams.set("scope", SCOPES);
			ms.searchParams.set("state", state);
			ms.searchParams.set("prompt", "select_account");

			return new Response(null, {
				status: 302,
				headers: {
					Location: ms.toString(),
					// Binds the flow to this browser (CSRF protection).
					"Set-Cookie": `${COOKIE}=${state}; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=600`,
				},
			});
		}

		if (url.pathname === "/callback") {
			const code = url.searchParams.get("code");
			const state = url.searchParams.get("state");
			const cookieState = request.headers.get("Cookie")?.match(new RegExp(`${COOKIE}=([^;]+)`))?.[1];
			if (!code || !state || state !== cookieState) {
				return new Response("Invalid or expired sign-in attempt. Try connecting again.", { status: 400 });
			}

			const stored = await env.OAUTH_KV.get(`ms_state:${state}`);
			if (!stored) return new Response("Sign-in expired. Try again.", { status: 400 });
			await env.OAUTH_KV.delete(`ms_state:${state}`); // single use
			const oauthReqInfo = JSON.parse(stored);

			const tokenRes = await fetch(`${MS_AUTH}/token`, {
				method: "POST",
				headers: { "Content-Type": "application/x-www-form-urlencoded" },
				body: new URLSearchParams({
					client_id: env.MICROSOFT_CLIENT_ID,
					client_secret: env.MICROSOFT_CLIENT_SECRET,
					grant_type: "authorization_code",
					code,
					redirect_uri: `${url.origin}/callback`,
					scope: SCOPES,
				}),
			});
			if (!tokenRes.ok) {
				return new Response(`Microsoft token exchange failed: ${await tokenRes.text()}`, { status: 502 });
			}
			const tok = (await tokenRes.json()) as {
				access_token: string;
				refresh_token: string;
				expires_in: number;
			};

			const meRes = await fetch(`${GRAPH}/me?$select=id,mail,userPrincipalName`, {
				headers: { Authorization: `Bearer ${tok.access_token}` },
			});
			if (!meRes.ok) return new Response("Could not read your Microsoft profile.", { status: 502 });
			const me = (await meRes.json()) as { id: string; mail?: string; userPrincipalName?: string };
			const email = (me.mail || me.userPrincipalName || "").toLowerCase();

			if (env.ALLOWED_EMAIL && email !== env.ALLOWED_EMAIL.toLowerCase()) {
				return new Response("This account is not allowed to use this server.", { status: 403 });
			}

			const { redirectTo } = await env.OAUTH_PROVIDER.completeAuthorization({
				request: oauthReqInfo,
				userId: me.id,
				metadata: { label: email },
				scope: oauthReqInfo.scope,
				props: {
					userId: me.id,
					email,
					refreshToken: tok.refresh_token,
					accessToken: tok.access_token,
					expiresAt: Date.now() + tok.expires_in * 1000,
				} satisfies Props,
			});

			return new Response(null, {
				status: 302,
				headers: {
					Location: redirectTo,
					"Set-Cookie": `${COOKIE}=; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=0`,
				},
			});
		}

		return new Response("Personal Microsoft MCP server. Connect via /mcp.", { status: 200 });
	},
};

export default new OAuthProvider({
	apiRoute: "/mcp",
	apiHandler: MicrosoftMCP.serve("/mcp") as any,
	defaultHandler: authHandler as any,
	authorizeEndpoint: "/authorize",
	tokenEndpoint: "/token",
	clientRegistrationEndpoint: "/register",
});
