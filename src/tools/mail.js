// Mail tools. Drafts only: deliberately no send tool (the user presses Send in Outlook).
import { z } from "zod";
import { request } from "../graph.js";

const recipients = (addresses = []) => addresses.map((address) => ({ emailAddress: { address } }));

export function register(server) {
  server.tool(
    "mail_create_draft",
    "Create a draft email in the Drafts folder of the signed-in mailbox. Does NOT send it.",
    {
      to: z.array(z.string().email()).min(1),
      subject: z.string(),
      body_html: z.string().describe("HTML body, e.g. <p>Hello</p>"),
      cc: z.array(z.string().email()).default([]),
      bcc: z.array(z.string().email()).default([]),
    },
    async ({ to, subject, body_html, cc, bcc }) => {
      const r = await request("POST", "me/messages", {
        json: {
          subject,
          body: { contentType: "HTML", content: body_html },
          toRecipients: recipients(to),
          ccRecipients: recipients(cc),
          bccRecipients: recipients(bcc),
        },
      });
      return {
        content: [
          { type: "text", text: JSON.stringify({ id: r.id, subject: r.subject, isDraft: r.isDraft, webLink: r.webLink }, null, 2) },
        ],
      };
    },
  );
}
