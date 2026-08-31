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
// NOTE: this is a drop-in replacement built from the documented behavior of
// the existing worker. If your deployed version has extra logic (custom CORS
// origin allow-list, rate limiting, etc.), port that over rather than
// pasting this verbatim — the important addition is the `action === 'delete'`
// branch below.

const REPO = 'FacilityAdmin-hub/karmsolar-fm-dashboard';

export default {
  async fetch(request, env) {
    const cors = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
    };

    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: cors });
    }
    if (request.method !== 'POST') {
      return new Response('Method not allowed', { status: 405, headers: cors });
    }

    let body;
    try {
      body = await request.json();
    } catch (e) {
      return new Response('Invalid JSON body', { status: 400, headers: cors });
    }

    let issueTitle, issueBody;

    if (body.action === 'delete') {
      const ticketId = String(body.ticketId || '').trim();
      if (!ticketId) {
        return new Response('Missing ticketId', { status: 400, headers: cors });
      }
      issueTitle = `Delete: ${ticketId}`;
      issueBody = JSON.stringify({ ticketId });
    } else {
      if (!body.title || !body.description) {
        return new Response('Missing required ticket fields', { status: 400, headers: cors });
      }
      issueTitle = `Ticket: ${body.title}`;
      issueBody = JSON.stringify(body);
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
      const text = await ghRes.text();
      return new Response(`GitHub issue creation failed: ${text}`, { status: 502, headers: cors });
    }

    return new Response('OK', { status: 200, headers: cors });
  },
};
