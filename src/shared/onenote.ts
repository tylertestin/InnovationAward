import { HostApp } from "./models";

export async function captureFromOneNotePage(): Promise<{
  host: HostApp;
  pageId?: string;
  title: string;
  extractedEmails: string[];
  extractedTextSample: string;
}> {
  try {
    const result = await OneNote.run(async (context) => {
      const page = context.application.getActivePageOrNull();
      page.load("id,title");

      // Load page contents (best effort).
      const contents = page.contents;
      contents.load("items/type,items/outline/paragraphs/type,items/outline/paragraphs/richText/text");

      await context.sync();

      if ((page as any).isNull) {
        return {
          pageId: undefined,
          title: "No active page",
          text: "",
        };
      }

      let text = "";
      try {
        const items = contents.items || [];
        for (const c of items) {
          if (c.type === "Outline" && c.outline) {
            const paragraphs = c.outline.paragraphs?.items || [];
            for (const p of paragraphs) {
              if (p.type === "RichText" && p.richText?.text) {
                text += p.richText.text + "\n";
              }
            }
          }
        }
      } catch {
        // ignore
      }

      return {
        pageId: page.id,
        title: page.title || "(untitled page)",
        text,
      };
    });

    const emailRegex = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;
    const emails = (result.text.match(emailRegex) || []).map((e) => e.toLowerCase());
    const deduped = Array.from(new Set(emails));

    const sample = result.text.trim().slice(0, 500);

    return {
      host: "OneNote",
      pageId: result.pageId,
      title: result.title,
      extractedEmails: deduped,
      extractedTextSample: sample,
    };
  } catch (e: any) {
    return {
      host: "OneNote",
      pageId: undefined,
      title: "Error reading OneNote page",
      extractedEmails: [],
      extractedTextSample: String(e?.message || e),
    };
  }
}
