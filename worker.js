// karmsolar-ticket-proxy — Cloudflare Worker
//
// Holds the real GitHub token as a private secret binding (env.GITHUB_TOKEN)
// and creates GitHub Issues on behalf of the public, sign-in-free pages:
//   - ticket-request.html  → { title, description, category, priority, site,
//                              requesterName, requesterEmail }
//                            → opens "Ticket: <title>" issue
//   - ticket-board.html    → { action: "delete", ticketId }
//                            → opens "Delete: <ticketId>" issue
//
// NOTE: this is a reference copy. The deployed worker also handles update-status,
// edit-ticket, add-comment, edit-comment and add-attachment. When updating the
// deployed worker, port the hardening below (origin allow-list, size limit, field
// allow-list, generic errors) into it rather than pasting this file verbatim.
// See SECURITY.md.
//
// Optional environment variables (Cloudflare dashboard → Worker → Settings → Variables):
//   ALLOWED_ORIGINS        comma-separated list of sites allowed to call this worker
//   ALLOWED_EMAIL_DOMAINS  comma-separated; requester emails must be on one of these (default karmsolar.com)

const REPO = 'FacilityAdmin-hub/karmsolar-fm-dashboard';

const DEFAULT_ORIGINS = [
  'https://facilityadmin-hub.github.io',
  'https://mohi899.github.io',
  'https://fm-dashboard-theta.vercel.app',
];

const MAX_BODY_BYTES = 32 * 1024;

const TICKET_FIELDS = {
  title: 200,
  description: 5000,
  category: 80,
  priority: 40,
  site: 120,
  requesterName: 120,
  requesterEmail: 254,
};

const TICKET_ID_RE = /^[A-Za-z0-9][A-Za-z0-9_-]{0,63}$/;

function list(value, fallback) {
  const items = String(value || '').split(',').map(s => s.trim()).filter(Boolean);
  return items.length ? items : fallback;
}

function clean(value, max) {
  // Strip control characters (keep newlines/tabs) and cap the length.
  return String(value == null ? '' : value).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').trim().slice(0, max);
}

export default {
  async fetch(request, env) {
    const allowedOrigins = list(env.ALLOWED_ORIGINS, DEFAULT_ORIGINS);
    const origin = request.headers.get('Origin') || '';
    const originOk = allowedOrigins.includes(origin);

    const cors = {
      'Access-Control-Allow-Origin': originOk ? origin : allowedOrigins[0],
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      'Access-Control-Max-Age': '86400',
      Vary: 'Origin',
      'X-Content-Type-Options': 'nosniff',
    };
    const reply = (text, status) => new Response(text, { status, headers: cors });

    if (request.method === 'OPTIONS') {
      return new Response(null, { status: originOk ? 204 : 403, headers: cors });
    }
    if (request.method !== 'POST') return reply('Method not allowed', 405);

    // Browsers always send Origin on cross-site POSTs. Requests from other sites
    // (or scripted requests without a matching Origin) are refused.
    if (!originOk) return reply('Forbidden', 403);

    const declared = Number(request.headers.get('Content-Length') || 0);
    if (declared > MAX_BODY_BYTES) return reply('Request too large', 413);

    let body;
    try {
      const raw = await request.text();
      if (raw.length > MAX_BODY_BYTES) return reply('Request too large', 413);
      body = JSON.parse(raw);
    } catch (e) {
      return reply('Invalid JSON body', 400);
    }
    if (!body || typeof body !== 'object' || Array.isArray(body)) return reply('Invalid request', 400);

    let issueTitle, issueBody;

    if (body.action === 'delete') {
      const ticketId = clean(body.ticketId, 64);
      if (!TICKET_ID_RE.test(ticketId)) return reply('Invalid ticketId', 400);
      issueTitle = `Delete: ${ticketId}`;
      issueBody = JSON.stringify({ ticketId });
    } else if (body.action === undefined || body.action === 'create') {
      // Copy only the known fields, so nothing unexpected ends up in the public issue.
      const ticket = {};
      for (const [key, max] of Object.entries(TICKET_FIELDS)) ticket[key] = clean(body[key], max);
      if (!ticket.title || !ticket.description) return reply('Missing required ticket fields', 400);

      const domains = list(env.ALLOWED_EMAIL_DOMAINS, ['karmsolar.com']).map(d => d.toLowerCase());
      const m = /^[^\s@<>"']+@([^\s@<>"']+)$/.exec(ticket.requesterEmail);
      if (ticket.requesterEmail && (!m || !domains.includes(m[1].toLowerCase()))) {
        return reply('Please use your company email address', 400);
      }

      issueTitle = `Ticket: ${ticket.title.replace(/\s+/g, ' ')}`.slice(0, 250);
      issueBody = JSON.stringify(ticket);
    } else {
      return reply('Unknown action', 400);
    }

    const ghRes = await fetch(`https://api.github.com/repos/${REPO}/issues`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${env.GITHUB_TOKEN}`,
        Accept: 'application/vnd.github+json',
        'Content-Type': 'application/json',
        'User-Agent': 'karmsolar-ticket-proxy',
      },
      body: JSON.stringify({ title: issueTitle, body: issueBody }),
    });

    if (!ghRes.ok) {
      // Log the details for the Cloudflare dashboard, but don't echo GitHub's response to the caller.
      console.log('GitHub issue creation failed', ghRes.status, await ghRes.text());
      return reply('Could not submit right now, please try again.', 502);
    }

    return reply('OK', 200);
  },
};
