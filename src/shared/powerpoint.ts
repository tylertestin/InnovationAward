export interface SlideCapture {
  slideId: string;
  slideText: string;
}

/**
 * Captures the current selected slide's text content by concatenating all shape text.
 * Requires the PowerPoint JavaScript API (Office.js) and a selected slide.
 */
export async function captureFromPowerPointSlide(): Promise<SlideCapture> {
  if (!(globalThis as any).PowerPoint?.run) {
    throw new Error("PowerPoint host not detected.");
  }

  return await (globalThis as any).PowerPoint.run(async (context: any) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.load("id");
    const shapes = slide.shapes;
    shapes.load("items/type,items/id,items/textFrame/textRange/text");
    await context.sync();

    const texts: string[] = [];
    for (const s of shapes.items) {
      const t = s?.textFrame?.textRange?.text;
      if (t && typeof t === "string" && t.trim().length > 0) texts.push(t.trim());
    }

    return {
      slideId: slide.id as string,
      slideText: texts.join("\n\n"),
    } as SlideCapture;
  });
}

export type StakeholderImpact = {
  stakeholderId: string;
  displayName: string;
  email?: string;
  reaction: "green" | "red";
  rationale?: string;
};
