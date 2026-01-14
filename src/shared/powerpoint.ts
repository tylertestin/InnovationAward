export interface SlideCapture {
  slideId: string;
  slideText: string;
}

async function ensureOfficeReady(): Promise<void> {
  const office = (globalThis as any).Office;
  if (office?.onReady) {
    await office.onReady();
  }
}

function isPowerPointApiSupported(): boolean {
  const office = (globalThis as any).Office;
  const requirements = office?.context?.requirements;
  if (!requirements?.isSetSupported) return true;
  return requirements.isSetSupported("PowerPointApi", "1.2");
}

async function captureWithPowerPointRun(powerpoint: any): Promise<SlideCapture> {
  return await powerpoint.run(async (context: any) => {
    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items/id");
    await context.sync();

    if (!selectedSlides.items.length) {
      throw new Error("No slide selected.");
    }

    const slide = selectedSlides.items[0];
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

async function captureWithOfficeSelection(): Promise<SlideCapture> {
  const office = (globalThis as any).Office;
  const documentApi = office?.context?.document;
  if (!documentApi?.getSelectedDataAsync) {
    throw new Error("PowerPoint host not detected.");
  }

  return await new Promise((resolve, reject) => {
    const coercionType = office.CoercionType?.SlideRange ?? office.CoercionType?.Text;
    documentApi.getSelectedDataAsync(
      coercionType,
      { valueFormat: office.ValueFormat?.Unformatted },
      (result: any) => {
        if (result?.status !== office.AsyncResultStatus?.Succeeded) {
          reject(new Error(result?.error?.message || "Unable to read selected slide."));
          return;
        }

        const value = result?.value;
        let slideId = "unknown";
        let slideText = "";

        if (typeof value === "string") {
          slideText = value;
        } else if (value?.slides?.length) {
          const slide = value.slides[0];
          slideId = slide?.id ?? slideId;
          if (Array.isArray(slide?.shapes)) {
            slideText = slide.shapes
              .map((shape: any) => {
                const text =
                  shape?.text ??
                  shape?.textRange?.text ??
                  shape?.textFrame?.textRange?.text ??
                  shape?.textFrame?.textRange?.text;
                return typeof text === "string" ? text.trim() : "";
              })
              .filter(Boolean)
              .join("\n\n");
          }
        } else if (typeof value?.text === "string") {
          slideText = value.text;
        }

        resolve({
          slideId,
          slideText,
        });
      }
    );
  });
}

/**
 * Captures the current selected slide's text content by concatenating all shape text.
 * Requires the PowerPoint JavaScript API (Office.js) and a selected slide.
 */
export async function captureFromPowerPointSlide(): Promise<SlideCapture> {
  await ensureOfficeReady();

  const powerpoint = (globalThis as any).PowerPoint;
  const canUsePowerPointApi = !!powerpoint?.run && isPowerPointApiSupported();

  if (canUsePowerPointApi) {
    try {
      return await captureWithPowerPointRun(powerpoint);
    } catch (error: any) {
      const office = (globalThis as any).Office;
      if (office?.context?.document?.getSelectedDataAsync) {
        try {
          return await captureWithOfficeSelection();
        } catch {
          throw error;
        }
      }
      throw error;
    }
  }

  return await captureWithOfficeSelection();
}

export type StakeholderImpact = {
  stakeholderId: string;
  displayName: string;
  email?: string;
  reaction: "green" | "red";
  rationale?: string;
};
