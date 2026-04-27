import { google, slides_v1 } from "googleapis";
import { OAuth2Client } from "google-auth-library";
import {
  ReviewData,
  SlideGenerationResult,
  SlideConfig,
  SlideElement,
  RgbColor,
} from "./types.js";
import { resolveColor } from "./config-loader.js";
import { getStatusEmoji } from "./templates.js";

// Points to EMU (English Metric Units) conversion
const PT_TO_EMU = 12700;
function pt(points: number): number {
  return points * PT_TO_EMU;
}

// Matrix animation constants
const MATRIX_CHARS =
  "ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝ";
const MATRIX_GREEN: RgbColor = { red: 0, green: 0.8, blue: 0.2 };
const FLASH_GREEN: RgbColor = { red: 0.3, green: 1.0, blue: 0.3 };
const PRESERVE_CHARS = new Set([
  ".", ",", "!", "?", ":", ";", "-", "'", '"', "(", ")", "[", "]",
]);
const RAIN_START_OFFSETS = [50, 100, 75]; // staggered heights per column (cycles)
const RAIN_Y_STEP = 25;            // points each rain column drops per frame
const RAIN_TAIL_FRAMES = 5;        // frames for text to settle after last deposit
const RAIN_FADE_STEPS = [3, 5, 4]; // rain drop fade-OUT steps per column (cycles)
const RAIN_FADE_IN_STEPS = [1, 3, 2]; // rain drop fade-IN steps per column (cycles)
const TEXT_FADE_STEPS = [4, 7, 5];  // text char fade-in steps: black→green per column
const RAIN_DROP_SIZE = 30;         // rain drop text box dimensions (points)
const CHAR_WIDTH_RATIO = 0.48;     // approximate char width / fontSize for Arial
const TEXT_BOX_PADDING_X = 7.2;    // text box internal left margin (~0.1in)

export function garbleText(text: string): string {
  return text
    .split("")
    .map((ch) => {
      if (ch === " " || PRESERVE_CHARS.has(ch)) return ch;
      return MATRIX_CHARS[Math.floor(Math.random() * MATRIX_CHARS.length)];
    })
    .join("");
}

/**
 * Generate slides from a SlideConfig
 */
export async function generateSlidesFromConfig(
  auth: OAuth2Client,
  config: SlideConfig
): Promise<SlideGenerationResult> {
  // Route to update-slide mode if specified
  if (config.presentationId && config.updateSlide !== undefined) {
    return updateSlideContent(auth, config);
  }

  const slidesApi = google.slides({ version: "v1", auth });
  const themeColors = config.theme?.colors;

  let presentationId = config.presentationId;
  let insertionIndexOffset = 0;

  if (presentationId && config.append) {
    // Append mode: keep existing slides, insert after them
    const existing = await slidesApi.presentations.get({ presentationId });
    const existingSlides = existing.data.slides || [];
    insertionIndexOffset = existingSlides.length;
  } else if (presentationId) {
    // Replace mode: delete all existing slides then rebuild
    const existing = await slidesApi.presentations.get({ presentationId });
    const existingSlides = existing.data.slides || [];

    if (existingSlides.length > 0) {
      const deleteRequests = existingSlides.map((slide) => ({
        deleteObject: { objectId: slide.objectId },
      }));
      await slidesApi.presentations.batchUpdate({
        presentationId,
        requestBody: { requests: deleteRequests },
      });
    }
  } else {
    // Create new presentation
    const presentation = await slidesApi.presentations.create({
      requestBody: { title: config.title },
    });
    presentationId = presentation.data.presentationId!;

    // Delete the default blank slide
    const defaultSlideId = presentation.data.slides![0].objectId;
    if (defaultSlideId) {
      await slidesApi.presentations.batchUpdate({
        presentationId,
        requestBody: {
          requests: [{ deleteObject: { objectId: defaultSlideId } }],
        },
      });
    }
  }

  // Build all requests
  const requests: slides_v1.Schema$Request[] = [];

  config.slides.forEach((slide, slideIndex) => {
    const slideId = `slide_${slideIndex + insertionIndexOffset}_${Date.now()}`;

    // Create slide with BLANK layout
    requests.push({
      createSlide: {
        objectId: slideId,
        insertionIndex: slideIndex + insertionIndexOffset,
        slideLayoutReference: { predefinedLayout: "BLANK" },
      },
    });

    // Set background if specified
    if (slide.background) {
      const bgColor = resolveColor(slide.background, themeColors);
      requests.push({
        updatePageProperties: {
          objectId: slideId,
          pageProperties: {
            pageBackgroundFill: {
              solidFill: { color: { rgbColor: bgColor } },
            },
          },
          fields: "pageBackgroundFill",
        },
      });
    }

    // Add text elements
    slide.elements.forEach((elem, elemIndex) => {
      const elementId = `${slideId}_text_${elemIndex}`;
      requests.push(...createTextBoxRequests(slideId, elementId, elem, themeColors));
    });
  });

  // Apply in batches (API limit is ~100 per request)
  const BATCH_SIZE = 50;
  for (let i = 0; i < requests.length; i += BATCH_SIZE) {
    const batch = requests.slice(i, i + BATCH_SIZE);
    await slidesApi.presentations.batchUpdate({
      presentationId,
      requestBody: { requests: batch },
    });
  }

  // Add speaker notes
  const notesRequests: slides_v1.Schema$Request[] = [];
  const presentation = await slidesApi.presentations.get({ presentationId });

  presentation.data.slides?.forEach((slide, i) => {
    const configIndex = i - insertionIndexOffset;
    if (configIndex < 0 || configIndex >= config.slides.length) return;
    if (!config.slides[configIndex].notes) return;

    const notesId =
      slide.slideProperties?.notesPage?.notesProperties?.speakerNotesObjectId;
    if (notesId) {
      notesRequests.push({
        insertText: {
          objectId: notesId,
          text: config.slides[configIndex].notes!,
          insertionIndex: 0,
        },
      });
    }
  });

  if (notesRequests.length > 0) {
    await slidesApi.presentations.batchUpdate({
      presentationId,
      requestBody: { requests: notesRequests },
    });
  }

  return {
    presentationId,
    presentationUrl: `https://docs.google.com/presentation/d/${presentationId}/edit`,
  };
}

/**
 * Update an existing slide's content in-place.
 * Deletes all elements on the target slide and rebuilds from config.slides[0].
 * Changes are visible in real-time to anyone viewing the presentation.
 */
async function updateSlideContent(
  auth: OAuth2Client,
  config: SlideConfig
): Promise<SlideGenerationResult> {
  const slidesApi = google.slides({ version: "v1", auth });
  const presentationId = config.presentationId!;
  const themeColors = config.theme?.colors;
  const slideTarget = config.updateSlide!;

  if (!config.slides.length) {
    throw new Error("No slides defined in config for update");
  }
  const slideDef = config.slides[0];

  // Fetch presentation to find the target slide
  const presentation = await slidesApi.presentations.get({ presentationId });
  const existingSlides = presentation.data.slides || [];

  if (existingSlides.length === 0) {
    throw new Error("Presentation has no slides to update");
  }

  // Resolve target index
  const targetIndex =
    slideTarget === "last" ? existingSlides.length - 1 : slideTarget;

  if (targetIndex < 0 || targetIndex >= existingSlides.length) {
    throw new Error(
      `Slide index ${targetIndex} out of range (0-${existingSlides.length - 1})`
    );
  }

  const targetSlide = existingSlides[targetIndex];
  const slideObjectId = targetSlide.objectId!;
  const requests: slides_v1.Schema$Request[] = [];

  // Delete all existing elements on the slide
  const elements = targetSlide.pageElements || [];
  for (const el of elements) {
    if (el.objectId) {
      requests.push({ deleteObject: { objectId: el.objectId } });
    }
  }

  // Update background if specified
  if (slideDef.background) {
    const bgColor = resolveColor(slideDef.background, themeColors);
    requests.push({
      updatePageProperties: {
        objectId: slideObjectId,
        pageProperties: {
          pageBackgroundFill: {
            solidFill: { color: { rgbColor: bgColor } },
          },
        },
        fields: "pageBackgroundFill",
      },
    });
  }

  // Add new elements, tracking any that need matrix animation
  const animatedElements: Array<{
    elementId: string;
    rainDropIds: string[];
    rainCharIndices: number[];
    originalText: string;
    fontSize: number;
    bold: boolean;
    elemX: number;
    elemY: number;
  }> = [];

  slideDef.elements.forEach((elem, elemIndex) => {
    const elementId = `${slideObjectId}_elem_${elemIndex}`;
    if (elem.animate === "matrix") {
      // One rain drop per non-space character position
      const rainCharIndices: number[] = [];
      for (let i = 0; i < elem.text.length; i++) {
        if (elem.text[i] !== " ") rainCharIndices.push(i);
      }
      const rainDropIds = rainCharIndices.map(
        (_, i) => `${elementId}_rain_${i}`
      );
      animatedElements.push({
        elementId,
        rainDropIds,
        rainCharIndices,
        originalText: elem.text,
        fontSize: elem.size,
        bold: elem.bold || false,
        elemX: elem.x,
        elemY: elem.y,
      });
      // Main text starts blank — rain drops are the only visual until they deposit
      const blankText = elem.text
        .split("")
        .map(ch => ch === " " ? ch : " ")
        .join("");
      requests.push(
        ...createTextBoxRequests(slideObjectId, elementId, elem, themeColors, {
          text: blankText,
          color: MATRIX_GREEN,
        })
      );
      // Create rain drops above the animated element (one per non-space character)
      const charWidth = elem.size * CHAR_WIDTH_RATIO;
      for (let i = 0; i < rainCharIndices.length; i++) {
        const charIndex = rainCharIndices[i];
        const dropX = elem.x + TEXT_BOX_PADDING_X + charIndex * charWidth;
        const dropY = elem.y - RAIN_START_OFFSETS[i % RAIN_START_OFFSETS.length];
        requests.push(
          {
            createShape: {
              objectId: rainDropIds[i],
              shapeType: "TEXT_BOX",
              elementProperties: {
                pageObjectId: slideObjectId,
                size: {
                  width: { magnitude: pt(RAIN_DROP_SIZE), unit: "EMU" },
                  height: { magnitude: pt(RAIN_DROP_SIZE), unit: "EMU" },
                },
                transform: {
                  scaleX: 1,
                  scaleY: 1,
                  translateX: pt(dropX),
                  translateY: pt(dropY),
                  unit: "EMU",
                },
              },
            },
          },
          {
            insertText: {
              objectId: rainDropIds[i],
              text: garbleText("X"),
            },
          },
          {
            updateTextStyle: {
              objectId: rainDropIds[i],
              style: {
                fontFamily: "Arial",
                fontSize: { magnitude: elem.size, unit: "PT" },
                foregroundColor: { opaqueColor: { rgbColor: { red: 0, green: 0, blue: 0 } } },
                bold: true,
              },
              fields: "fontFamily,fontSize,foregroundColor,bold",
            },
          }
        );
      }
    } else {
      requests.push(
        ...createTextBoxRequests(slideObjectId, elementId, elem, themeColors)
      );
    }
  });

  // Apply element requests in batches
  const BATCH_SIZE = 50;
  for (let i = 0; i < requests.length; i += BATCH_SIZE) {
    const batch = requests.slice(i, i + BATCH_SIZE);
    await slidesApi.presentations.batchUpdate({
      presentationId,
      requestBody: { requests: batch },
    });
  }

  // Run matrix reveal animation for animated elements
  for (const anim of animatedElements) {
    await animateMatrixReveal(
      slidesApi,
      presentationId,
      anim.elementId,
      anim.rainDropIds,
      anim.rainCharIndices,
      anim.originalText,
      anim.fontSize,
      anim.bold,
      anim.elemX,
      anim.elemY
    );
  }

  // Update speaker notes (clear existing, then insert new)
  if (slideDef.notes) {
    const notesId =
      targetSlide.slideProperties?.notesPage?.notesProperties
        ?.speakerNotesObjectId;
    if (notesId) {
      // Check if notes have existing text before trying to delete
      const notesPage = targetSlide.slideProperties?.notesPage;
      const notesElement = notesPage?.pageElements?.find(
        (el) => el.objectId === notesId
      );
      const hasExistingText =
        notesElement?.shape?.text?.textElements?.some(
          (te) => te.textRun?.content && te.textRun.content.trim().length > 0
        ) ?? false;

      const requests: slides_v1.Schema$Request[] = [];
      if (hasExistingText) {
        requests.push({
          deleteText: {
            objectId: notesId,
            textRange: { type: "ALL" },
          },
        });
      }
      requests.push({
        insertText: {
          objectId: notesId,
          text: slideDef.notes,
          insertionIndex: 0,
        },
      });

      await slidesApi.presentations.batchUpdate({
        presentationId,
        requestBody: { requests },
      });
    }
  }

  return {
    presentationId,
    presentationUrl: `https://docs.google.com/presentation/d/${presentationId}/edit`,
  };
}

async function animateMatrixReveal(
  slidesApi: slides_v1.Slides,
  presentationId: string,
  elementId: string,
  rainDropIds: string[],
  rainCharIndices: number[],
  originalText: string,
  fontSize: number,
  bold: boolean,
  elemX: number,
  elemY: number
): Promise<void> {
  const baseStyle = {
    fontFamily: "Arial",
    fontSize: { magnitude: fontSize, unit: "PT" } as const,
    bold,
  };

  const charWidth = fontSize * CHAR_WIDTH_RATIO;

  // Each column gets a staggered start offset (cycles through the array)
  const startOffsets = rainDropIds.map(
    (_, i) => RAIN_START_OFFSETS[i % RAIN_START_OFFSETS.length]
  );
  const maxStartOffset = Math.max(...startOffsets, RAIN_START_OFFSETS[0]);
  const lastDepositFrame = Math.ceil(maxStartOffset / RAIN_Y_STEP) - 1;
  const totalFrames = lastDepositFrame + 1 + RAIN_TAIL_FRAMES + 1;

  // Reverse map: char index → rain drop index (for per-char fade timing)
  const charToRainIndex = new Map<number, number>();
  for (let i = 0; i < rainCharIndices.length; i++) {
    charToRainIndex.set(rainCharIndices[i], i);
  }

  for (let frame = 0; frame < totalFrames; frame++) {
    const requests: slides_v1.Schema$Request[] = [];
    const isFinalFrame = frame === totalFrames - 1;

    // Determine which characters are deposited on this frame
    const depositedIndices = new Set<number>();
    for (let i = 0; i < rainDropIds.length; i++) {
      const offset = startOffsets[i];
      const charIdx = rainCharIndices[i];
      if ((frame + 1) * RAIN_Y_STEP >= offset) {
        depositedIndices.add(charIdx);
      }
    }

    // === Update main line text ===
    requests.push({
      deleteText: {
        objectId: elementId,
        textRange: { type: "ALL" },
      },
    });

    if (depositedIndices.size > 0) {
      // Build text: deposited chars show original, rest are blank
      const frameText = originalText
        .split("")
        .map((ch, i) => {
          if (ch === " ") return ch;
          if (depositedIndices.has(i)) return ch;
          return " ";
        })
        .join("");
      requests.push({
        insertText: {
          objectId: elementId,
          text: frameText,
        },
      });
      // Base style: black (invisible baseline)
      requests.push({
        updateTextStyle: {
          objectId: elementId,
          style: {
            ...baseStyle,
            foregroundColor: { opaqueColor: { rgbColor: { red: 0, green: 0, blue: 0 } } },
          },
          textRange: {
            type: "FIXED_RANGE",
            startIndex: 0,
            endIndex: originalText.length,
          },
          fields: "fontFamily,fontSize,foregroundColor,bold",
        },
      });
      // Per-character fade: black → FLASH_GREEN (stays green, no white transition)
      for (const charIdx of depositedIndices) {
        const rainIdx = charToRainIndex.get(charIdx)!;
        const offset = startOffsets[rainIdx];
        const depositFrame = Math.ceil(offset / RAIN_Y_STEP) - 1;
        const framesSinceDeposit = frame - depositFrame;
        const textFadeTotal = TEXT_FADE_STEPS[rainIdx % TEXT_FADE_STEPS.length];
        const fadeProgress = Math.min(framesSinceDeposit / textFadeTotal, 1.0);
        const color = {
          red: FLASH_GREEN.red * fadeProgress,
          green: FLASH_GREEN.green * fadeProgress,
          blue: FLASH_GREEN.blue * fadeProgress,
        };
        requests.push({
          updateTextStyle: {
            objectId: elementId,
            style: {
              ...baseStyle,
              foregroundColor: { opaqueColor: { rgbColor: color } },
            },
            textRange: {
              type: "FIXED_RANGE",
              startIndex: charIdx,
              endIndex: charIdx + 1,
            },
            fields: "fontFamily,fontSize,foregroundColor,bold",
          },
        });
      }
    } else {
      // No deposits yet — blank text (rain drops are the only visual)
      const blankText = originalText
        .split("")
        .map(ch => ch === " " ? ch : " ")
        .join("");
      requests.push(
        {
          insertText: {
            objectId: elementId,
            text: blankText,
          },
        },
        {
          updateTextStyle: {
            objectId: elementId,
            style: {
              ...baseStyle,
              foregroundColor: { opaqueColor: { rgbColor: { red: 0, green: 0, blue: 0 } } },
            },
            textRange: {
              type: "FIXED_RANGE",
              startIndex: 0,
              endIndex: originalText.length,
            },
            fields: "fontFamily,fontSize,foregroundColor,bold",
          },
        }
      );
    }

    // === Update rain drops ===
    if (isFinalFrame) {
      // Cleanup: delete all rain drops
      for (const id of rainDropIds) {
        requests.push({ deleteObject: { objectId: id } });
      }
    } else {
      // Move each rain drop; fade in before deposit, fade out after
      for (let i = 0; i < rainDropIds.length; i++) {
        const charIndex = rainCharIndices[i];
        const dropX = elemX + TEXT_BOX_PADDING_X + charIndex * charWidth;
        const offset = startOffsets[i];
        const rainY = elemY - offset + (frame + 1) * RAIN_Y_STEP;
        const hasDeposited = (frame + 1) * RAIN_Y_STEP >= offset;
        let dropColor: RgbColor;
        if (!hasDeposited) {
          // Fading in: black → FLASH_GREEN at different rates
          const fadeInSteps = RAIN_FADE_IN_STEPS[i % RAIN_FADE_IN_STEPS.length];
          const fadeInProgress = Math.min((frame + 1) / fadeInSteps, 1.0);
          dropColor = {
            red: FLASH_GREEN.red * fadeInProgress,
            green: FLASH_GREEN.green * fadeInProgress,
            blue: FLASH_GREEN.blue * fadeInProgress,
          };
        } else {
          const depositFrame = Math.ceil(offset / RAIN_Y_STEP) - 1;
          const framesSinceDeposit = frame - depositFrame;
          if (framesSinceDeposit === 0) {
            dropColor = FLASH_GREEN;
          } else {
            const fadeSteps = RAIN_FADE_STEPS[i % RAIN_FADE_STEPS.length];
            const fadeProgress = Math.min(framesSinceDeposit / fadeSteps, 1.0);
            dropColor = {
              red: MATRIX_GREEN.red * (1 - fadeProgress),
              green: MATRIX_GREEN.green * (1 - fadeProgress),
              blue: MATRIX_GREEN.blue * (1 - fadeProgress),
            };
          }
        }
        requests.push(
          { deleteText: { objectId: rainDropIds[i], textRange: { type: "ALL" } } },
          { insertText: { objectId: rainDropIds[i], text: garbleText("X") } },
          {
            updateTextStyle: {
              objectId: rainDropIds[i],
              style: {
                fontFamily: "Arial",
                fontSize: { magnitude: fontSize, unit: "PT" },
                foregroundColor: { opaqueColor: { rgbColor: dropColor } },
                bold: true,
              },
              fields: "fontFamily,fontSize,foregroundColor,bold",
            },
          },
          {
            updatePageElementTransform: {
              objectId: rainDropIds[i],
              applyMode: "ABSOLUTE",
              transform: {
                scaleX: 1,
                scaleY: 1,
                translateX: pt(dropX),
                translateY: pt(rainY),
                unit: "EMU",
              },
            },
          }
        );
      }
    }

    // Each batchUpdate is a visible "frame" — API latency provides natural pacing
    await slidesApi.presentations.batchUpdate({
      presentationId,
      requestBody: { requests },
    });
  }
}

function createTextBoxRequests(
  slideId: string,
  elementId: string,
  elem: SlideElement,
  themeColors?: Record<string, RgbColor>,
  overrides?: { text?: string; color?: RgbColor }
): slides_v1.Schema$Request[] {
  const color = overrides?.color || resolveColor(elem.color, themeColors);
  const text = overrides?.text ?? elem.text;

  return [
    {
      createShape: {
        objectId: elementId,
        shapeType: "TEXT_BOX",
        elementProperties: {
          pageObjectId: slideId,
          size: {
            width: { magnitude: pt(elem.w), unit: "EMU" },
            height: { magnitude: pt(elem.h), unit: "EMU" },
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: pt(elem.x),
            translateY: pt(elem.y),
            unit: "EMU",
          },
        },
      },
    },
    {
      insertText: {
        objectId: elementId,
        text: text,
      },
    },
    {
      updateTextStyle: {
        objectId: elementId,
        style: {
          fontFamily: "Arial",
          fontSize: { magnitude: elem.size, unit: "PT" },
          foregroundColor: { opaqueColor: { rgbColor: color } },
          bold: elem.bold || false,
        },
        fields: "fontFamily,fontSize,foregroundColor,bold",
      },
    },
    {
      updateParagraphStyle: {
        objectId: elementId,
        style: {
          lineSpacing: 115,
          alignment: "START",
        },
        fields: "lineSpacing,alignment",
      },
    },
  ];
}

/**
 * Legacy function: Generate slides from ReviewData
 * Kept for backward compatibility
 */
export async function generateSlides(
  auth: OAuth2Client,
  reviewData: ReviewData
): Promise<SlideGenerationResult> {
  const slides = google.slides({ version: "v1", auth });

  const presentation = await slides.presentations.create({
    requestBody: {
      title: `PR Review: ${reviewData.prTitle}`,
    },
  });

  const presentationId = presentation.data.presentationId!;
  const defaultSlideId = presentation.data.slides![0].objectId!;

  const slideIds = {
    title: `title_${Date.now()}`,
    summary: `summary_${Date.now()}`,
    impact: `impact_${Date.now()}`,
    risks: `risks_${Date.now()}`,
    verdict: `verdict_${Date.now()}`,
  };

  const createRequests: slides_v1.Schema$Request[] = [
    { deleteObject: { objectId: defaultSlideId } },
    ...Object.values(slideIds).map((id) => ({
      createSlide: {
        objectId: id,
        slideLayoutReference: { predefinedLayout: "TITLE_AND_BODY" as const },
      },
    })),
  ];

  await slides.presentations.batchUpdate({
    presentationId,
    requestBody: { requests: createRequests },
  });

  const updatedPresentation = await slides.presentations.get({
    presentationId,
  });

  const contentRequests: slides_v1.Schema$Request[] = [];

  for (const slide of updatedPresentation.data.slides || []) {
    const slideId = slide.objectId!;
    const placeholders = slide.pageElements?.filter(
      (el) => el.shape?.placeholder
    );

    const titlePlaceholder = placeholders?.find(
      (el) => el.shape?.placeholder?.type === "TITLE"
    );
    const bodyPlaceholder = placeholders?.find(
      (el) => el.shape?.placeholder?.type === "BODY"
    );

    const titleId = titlePlaceholder?.objectId ?? undefined;
    const bodyId = bodyPlaceholder?.objectId ?? undefined;

    if (slideId === slideIds.title) {
      contentRequests.push(...buildTitleContent(titleId, bodyId, reviewData));
    } else if (slideId === slideIds.summary) {
      contentRequests.push(...buildSummaryContent(titleId, bodyId, reviewData));
    } else if (slideId === slideIds.impact) {
      contentRequests.push(...buildImpactContent(titleId, bodyId, reviewData));
    } else if (slideId === slideIds.risks) {
      contentRequests.push(...buildRisksContent(titleId, bodyId, reviewData));
    } else if (slideId === slideIds.verdict) {
      contentRequests.push(...buildVerdictContent(titleId, bodyId, reviewData));
    }
  }

  if (contentRequests.length > 0) {
    await slides.presentations.batchUpdate({
      presentationId,
      requestBody: { requests: contentRequests },
    });
  }

  return {
    presentationId,
    presentationUrl: `https://docs.google.com/presentation/d/${presentationId}/edit`,
  };
}

function buildTitleContent(
  titleId: string | undefined,
  bodyId: string | undefined,
  data: ReviewData
): slides_v1.Schema$Request[] {
  const requests: slides_v1.Schema$Request[] = [];

  if (titleId) {
    requests.push({
      insertText: {
        objectId: titleId,
        text: data.prTitle,
        insertionIndex: 0,
      },
    });
  }

  if (bodyId) {
    const subtitle = `PR #${data.prNumber} | ${data.repository}\n${data.prAuthor} | ${formatDate(data.prDate)}`;
    requests.push({
      insertText: {
        objectId: bodyId,
        text: subtitle,
        insertionIndex: 0,
      },
    });
  }

  return requests;
}

function buildSummaryContent(
  titleId: string | undefined,
  bodyId: string | undefined,
  data: ReviewData
): slides_v1.Schema$Request[] {
  const requests: slides_v1.Schema$Request[] = [];

  if (titleId) {
    requests.push({
      insertText: {
        objectId: titleId,
        text: "What Changed",
        insertionIndex: 0,
      },
    });
  }

  if (bodyId) {
    const lines = [data.summary, "", ...data.changes.map((c) => `- ${c}`)];
    requests.push({
      insertText: {
        objectId: bodyId,
        text: lines.join("\n"),
        insertionIndex: 0,
      },
    });
  }

  return requests;
}

function buildImpactContent(
  titleId: string | undefined,
  bodyId: string | undefined,
  data: ReviewData
): slides_v1.Schema$Request[] {
  const requests: slides_v1.Schema$Request[] = [];

  if (titleId) {
    requests.push({
      insertText: {
        objectId: titleId,
        text: "Impact Assessment",
        insertionIndex: 0,
      },
    });
  }

  if (bodyId) {
    const lines: string[] = [];

    if (data.businessImpact) {
      lines.push("Business Impact:", data.businessImpact, "");
    }

    if (data.affectedAreas && data.affectedAreas.length > 0) {
      lines.push("Affected Areas:");
      lines.push(...data.affectedAreas.map((a) => `- ${a}`));
      lines.push("");
    }

    lines.push("Quality Summary:");
    const qa = data.qualityAssessment;
    lines.push(`- Code Quality: ${getStatusEmoji(qa.codeQuality.status)}`);
    lines.push(`- Tests: ${getStatusEmoji(qa.tests.status)}`);
    lines.push(`- Security: ${getStatusEmoji(qa.security.status)}`);
    lines.push(`- Performance: ${getStatusEmoji(qa.performance.status)}`);

    requests.push({
      insertText: {
        objectId: bodyId,
        text: lines.join("\n"),
        insertionIndex: 0,
      },
    });
  }

  return requests;
}

function buildRisksContent(
  titleId: string | undefined,
  bodyId: string | undefined,
  data: ReviewData
): slides_v1.Schema$Request[] {
  const requests: slides_v1.Schema$Request[] = [];

  if (titleId) {
    requests.push({
      insertText: {
        objectId: titleId,
        text: "Risk Assessment",
        insertionIndex: 0,
      },
    });
  }

  if (bodyId) {
    const lines: string[] = [];

    const riskLevel = data.riskLevel || "low";
    lines.push(`Risk Level: ${riskLevel.toUpperCase()}`, "");

    if (data.riskFactors && data.riskFactors.length > 0) {
      lines.push("Risk Factors:");
      lines.push(...data.riskFactors.map((r) => `- ${r}`));
      lines.push("");
    }

    if (data.issuesFound && data.issuesFound.length > 0) {
      lines.push("Issues Found:");
      lines.push(...data.issuesFound.map((i) => `- ${i}`));
    } else {
      lines.push("No blocking issues found.");
    }

    requests.push({
      insertText: {
        objectId: bodyId,
        text: lines.join("\n"),
        insertionIndex: 0,
      },
    });
  }

  return requests;
}

function buildVerdictContent(
  titleId: string | undefined,
  bodyId: string | undefined,
  data: ReviewData
): slides_v1.Schema$Request[] {
  const requests: slides_v1.Schema$Request[] = [];

  if (titleId) {
    requests.push({
      insertText: {
        objectId: titleId,
        text: `Recommendation: ${data.verdict}`,
        insertionIndex: 0,
      },
    });
  }

  if (bodyId) {
    const lines: string[] = [data.verdictExplanation, ""];

    if (data.suggestions && data.suggestions.length > 0) {
      lines.push("Suggestions:");
      lines.push(...data.suggestions.map((s) => `- ${s}`));
    }

    requests.push({
      insertText: {
        objectId: bodyId,
        text: lines.join("\n"),
        insertionIndex: 0,
      },
    });
  }

  return requests;
}

function formatDate(isoDate: string): string {
  return new Date(isoDate).toLocaleDateString("en-US", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });
}
