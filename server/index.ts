import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import OpenAI from "openai";

dotenv.config();

const app = express();
app.use(cors({ origin: true, credentials: true }));
app.use(express.json({ limit: "2mb" }));

const openaiKey = process.env.OPENAI_API_KEY;
if (!openaiKey) {
  // Server can still start, but OpenAI routes will fail.
  console.warn("OPENAI_API_KEY not set. Set it in .env for /api/openai/* routes.");
}
const openai = new OpenAI({ apiKey: openaiKey });

// Note: Outlook ingestion is handled via static user exports (CSV) in the client.
// We intentionally do not connect to Microsoft Graph here, to avoid tenant app-registration requirements.

/**
 * POST /api/openai/stakeholder-impact
 * Body: {
 *   stakeholders: Array<{id, displayName, email?, notes?: any, tags?: string[]}>,
 *   slideText?: string,
 *   emails?: Array<{subject, from, bodyPreview, receivedDateTime}>,
 * }
 *
 * Returns: { impacts: Array<{ stakeholderId, reaction, rationale }> }
 */
app.post("/api/openai/stakeholder-impact", async (req, res) => {
  try {
    if (!openaiKey) return res.status(500).json({ error: "OPENAI_API_KEY missing on server" });

    const { stakeholders, slideText, emails } = req.body ?? {};
    if (!Array.isArray(stakeholders)) return res.status(400).json({ error: "stakeholders must be an array" });

    const prompt = {
      role: "user" as const,
      content: [
        {
          type: "text" as const,
          text:
`You are assisting a consulting team. Given (1) a slide's content and (2) recent inbox context, identify which stakeholders are impacted and predict their reaction.
Return STRICT JSON: { "impacts": [ { "stakeholderId": "...", "reaction": "green"|"red", "rationale": "..." } ] }.
Definitions:
- green = likely supportive / aligned / positive
- red = likely skeptical / concerned / negative or risk
Only include stakeholders that are meaningfully impacted by the content.
Use concise rationales (<= 20 words).`,
        },
        { type: "text" as const, text: `SLIDE_TEXT:\n${String(slideText ?? "").slice(0, 12000)}` },
        { type: "text" as const, text: `STAKEHOLDERS:\n${JSON.stringify(stakeholders).slice(0, 12000)}` },
        { type: "text" as const, text: `RECENT_EMAILS:\n${JSON.stringify(emails ?? []).slice(0, 12000)}` },
      ],
    };

    const completion = await openai.chat.completions.create({
      model: process.env.OPENAI_MODEL || "gpt-4.1-mini",
      messages: [prompt],
      temperature: 0.2,
      response_format: { type: "json_object" },
    });

    const txt = completion.choices?.[0]?.message?.content ?? "{}";
    let parsed: any;
    try {
      parsed = JSON.parse(txt);
    } catch {
      return res.status(200).json({ impacts: [], raw: txt });
    }
    return res.status(200).json(parsed);
  } catch (e: any) {
    console.error(e);
    return res.status(500).json({ error: e?.message || "server error" });
  }
});

app.get("/health", (_req, res) => res.json({ ok: true }));

const port = Number(process.env.API_PORT || 3001);
app.listen(port, () => console.log(`API listening on http://localhost:${port}`));
