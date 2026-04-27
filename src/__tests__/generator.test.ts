import { describe, it, expect, vi, beforeEach } from "vitest";
import {
  generateSlidesFromConfig,
  generateSlides,
  garbleText,
} from "../slides/generator.js";
import { SlideConfig, ReviewData } from "../slides/types.js";
import { google } from "googleapis";
import { OAuth2Client } from "google-auth-library";

vi.mock("googleapis", () => {
  const mockBatchUpdate = vi.fn().mockResolvedValue({});
  const mockGet = vi.fn();
  const mockCreate = vi.fn();

  return {
    google: {
      slides: () => ({
        presentations: {
          get: mockGet,
          create: mockCreate,
          batchUpdate: mockBatchUpdate,
        },
      }),
    },
  };
});

function getMocks() {
  const api = google.slides({ version: "v1" });
  return {
    get: api.presentations.get as ReturnType<typeof vi.fn>,
    create: api.presentations.create as ReturnType<typeof vi.fn>,
    batchUpdate: api.presentations.batchUpdate as ReturnType<typeof vi.fn>,
  };
}

const mockAuth = {} as OAuth2Client;

function makeSlide(objectId: string, pageElements?: { objectId: string }[]) {
  return {
    objectId,
    pageElements: pageElements || [],
    slideProperties: {
      notesPage: {
        notesProperties: {
          speakerNotesObjectId: `notes_${objectId}`,
        },
      },
    },
  };
}

function makeLegacySlide(objectId: string, titleId: string, bodyId: string) {
  return {
    objectId,
    pageElements: [
      {
        objectId: titleId,
        shape: { placeholder: { type: "TITLE" } },
      },
      {
        objectId: bodyId,
        shape: { placeholder: { type: "BODY" } },
      },
    ],
  };
}

// Helper to collect all requests from all batchUpdate calls
function getAllRequests(batchUpdate: ReturnType<typeof vi.fn>) {
  const all: Record<string, unknown>[] = [];
  for (const call of batchUpdate.mock.calls) {
    const requests = call[0]?.requestBody?.requests || [];
    all.push(...requests);
  }
  return all;
}

// Helper to get requests from last batchUpdate call
function getLastRequests(batchUpdate: ReturnType<typeof vi.fn>) {
  const lastCall = batchUpdate.mock.calls[batchUpdate.mock.calls.length - 1];
  return lastCall?.[0]?.requestBody?.requests || [];
}

describe("generateSlidesFromConfig", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  const singleSlideConfig: SlideConfig = {
    title: "Test",
    slides: [
      {
        elements: [
          {
            text: "Hello",
            x: 0,
            y: 0,
            w: 100,
            h: 50,
            size: 18,
            color: "white",
          },
        ],
        notes: "Speaker note",
      },
    ],
  };

  describe("append mode", () => {
    it("should NOT delete existing slides", async () => {
      const { get, batchUpdate } = getMocks();

      get.mockResolvedValueOnce({
        data: {
          slides: [
            makeSlide("existing_1"),
            makeSlide("existing_2"),
            makeSlide("existing_3"),
          ],
        },
      }).mockResolvedValueOnce({
        data: {
          slides: [
            makeSlide("existing_1"),
            makeSlide("existing_2"),
            makeSlide("existing_3"),
            makeSlide("new_slide"),
          ],
        },
      });

      const config: SlideConfig = {
        ...singleSlideConfig,
        presentationId: "test-pres-id",
        append: true,
      };

      await generateSlidesFromConfig(mockAuth, config);

      const allReqs = getAllRequests(batchUpdate);
      for (const req of allReqs) {
        expect(req).not.toHaveProperty("deleteObject");
      }
    });

    it("should use correct insertionIndex offset", async () => {
      const { get, batchUpdate } = getMocks();

      get.mockResolvedValueOnce({
        data: {
          slides: [
            makeSlide("existing_1"),
            makeSlide("existing_2"),
            makeSlide("existing_3"),
          ],
        },
      }).mockResolvedValueOnce({
        data: {
          slides: [
            makeSlide("existing_1"),
            makeSlide("existing_2"),
            makeSlide("existing_3"),
            makeSlide("new_slide"),
          ],
        },
      });

      const config: SlideConfig = {
        ...singleSlideConfig,
        presentationId: "test-pres-id",
        append: true,
      };

      await generateSlidesFromConfig(mockAuth, config);

      const allReqs = getAllRequests(batchUpdate);
      const createSlideReq = allReqs.find(
        (r) => r.createSlide
      )?.createSlide as Record<string, unknown>;

      expect(createSlideReq).toBeDefined();
      expect(createSlideReq.insertionIndex).toBe(3);
    });

    it("should target speaker notes at appended slides only", async () => {
      const { get, batchUpdate } = getMocks();

      get.mockResolvedValueOnce({
        data: {
          slides: [makeSlide("existing_1"), makeSlide("existing_2")],
        },
      }).mockResolvedValueOnce({
        data: {
          slides: [
            makeSlide("existing_1"),
            makeSlide("existing_2"),
            makeSlide("appended_slide"),
          ],
        },
      });

      const config: SlideConfig = {
        ...singleSlideConfig,
        presentationId: "test-pres-id",
        append: true,
      };

      await generateSlidesFromConfig(mockAuth, config);

      const lastReqs = getLastRequests(batchUpdate);
      const notesReqs = lastReqs.filter(
        (r: Record<string, unknown>) =>
          r.insertText &&
          (r.insertText as Record<string, unknown>).objectId ===
            "notes_appended_slide"
      );

      expect(notesReqs.length).toBe(1);
      expect(
        (notesReqs[0].insertText as Record<string, unknown>).text
      ).toBe("Speaker note");
    });

    it("should handle appending to empty presentation", async () => {
      const { get, batchUpdate } = getMocks();

      get.mockResolvedValueOnce({
        data: { slides: [] },
      }).mockResolvedValueOnce({
        data: { slides: [makeSlide("new_slide")] },
      });

      const config: SlideConfig = {
        ...singleSlideConfig,
        presentationId: "test-pres-id",
        append: true,
      };

      await generateSlidesFromConfig(mockAuth, config);

      const allReqs = getAllRequests(batchUpdate);
      const createSlideReq = allReqs.find(
        (r) => r.createSlide
      )?.createSlide as Record<string, unknown>;

      expect(createSlideReq.insertionIndex).toBe(0);
    });
  });

  describe("replace mode (existing behavior)", () => {
    it("should delete existing slides before rebuilding", async () => {
      const { get, batchUpdate } = getMocks();

      get.mockResolvedValueOnce({
        data: {
          slides: [makeSlide("old_1"), makeSlide("old_2")],
        },
      }).mockResolvedValueOnce({
        data: {
          slides: [makeSlide("new_slide")],
        },
      });

      const config: SlideConfig = {
        ...singleSlideConfig,
        presentationId: "test-pres-id",
      };

      await generateSlidesFromConfig(mockAuth, config);

      const firstCall = batchUpdate.mock.calls[0];
      const deleteReqs = (firstCall[0]?.requestBody?.requests || []).filter(
        (r: Record<string, unknown>) => r.deleteObject
      );

      expect(deleteReqs.length).toBe(2);
    });

    it("should handle replacing with no existing slides", async () => {
      const { get, batchUpdate } = getMocks();

      get.mockResolvedValueOnce({
        data: { slides: [] },
      }).mockResolvedValueOnce({
        data: { slides: [makeSlide("new_slide")] },
      });

      const config: SlideConfig = {
        ...singleSlideConfig,
        presentationId: "test-pres-id",
      };

      await generateSlidesFromConfig(mockAuth, config);

      // First batchUpdate should be the createSlide, not a delete
      const firstReqs = batchUpdate.mock.calls[0][0]?.requestBody?.requests || [];
      const hasDelete = firstReqs.some(
        (r: Record<string, unknown>) => r.deleteObject
      );
      expect(hasDelete).toBe(false);
    });
  });

  describe("new presentation mode", () => {
    it("should create a new presentation and delete default slide", async () => {
      const { create, get, batchUpdate } = getMocks();

      create.mockResolvedValue({
        data: {
          presentationId: "new-pres-id",
          slides: [{ objectId: "default_slide" }],
        },
      });

      get.mockResolvedValueOnce({
        data: {
          slides: [makeSlide("created_slide")],
        },
      });

      const config: SlideConfig = { ...singleSlideConfig };

      const result = await generateSlidesFromConfig(mockAuth, config);

      expect(create).toHaveBeenCalledWith({
        requestBody: { title: "Test" },
      });
      expect(result.presentationId).toBe("new-pres-id");
      expect(result.presentationUrl).toContain("new-pres-id");

      // Should delete the default blank slide
      const firstBatchReqs =
        batchUpdate.mock.calls[0][0]?.requestBody?.requests || [];
      expect(firstBatchReqs[0]).toEqual({
        deleteObject: { objectId: "default_slide" },
      });
    });
  });

  describe("slide content generation", () => {
    it("should create text box with correct EMU dimensions", async () => {
      const { create, get, batchUpdate } = getMocks();

      create.mockResolvedValue({
        data: {
          presentationId: "pres-id",
          slides: [{ objectId: "default" }],
        },
      });

      get.mockResolvedValueOnce({
        data: { slides: [makeSlide("s")] },
      });

      const config: SlideConfig = {
        title: "Test",
        slides: [
          {
            elements: [
              {
                text: "Hello",
                x: 10,
                y: 20,
                w: 100,
                h: 50,
                size: 18,
                color: "white",
              },
            ],
          },
        ],
      };

      await generateSlidesFromConfig(mockAuth, config);

      const allReqs = getAllRequests(batchUpdate);
      const createShape = allReqs.find(
        (r) => r.createShape
      )?.createShape as Record<string, unknown>;

      expect(createShape).toBeDefined();
      const elemProps = createShape.elementProperties as Record<string, unknown>;
      const size = elemProps.size as Record<string, unknown>;
      const width = size.width as Record<string, unknown>;
      const height = size.height as Record<string, unknown>;

      // 100 points * 12700 EMU/pt = 1,270,000 EMU
      expect(width.magnitude).toBe(100 * 12700);
      expect(height.magnitude).toBe(50 * 12700);
      expect(width.unit).toBe("EMU");
    });

    it("should set text style with font size and bold", async () => {
      const { create, get, batchUpdate } = getMocks();

      create.mockResolvedValue({
        data: {
          presentationId: "pres-id",
          slides: [{ objectId: "default" }],
        },
      });

      get.mockResolvedValueOnce({
        data: { slides: [makeSlide("s")] },
      });

      const config: SlideConfig = {
        title: "Test",
        slides: [
          {
            elements: [
              {
                text: "Bold text",
                x: 0,
                y: 0,
                w: 100,
                h: 50,
                size: 24,
                color: "white",
                bold: true,
              },
            ],
          },
        ],
      };

      await generateSlidesFromConfig(mockAuth, config);

      const allReqs = getAllRequests(batchUpdate);
      const textStyle = allReqs.find(
        (r) => r.updateTextStyle
      )?.updateTextStyle as Record<string, unknown>;

      expect(textStyle).toBeDefined();
      const style = textStyle.style as Record<string, unknown>;
      expect(style.bold).toBe(true);
      const fontSize = style.fontSize as Record<string, unknown>;
      expect(fontSize.magnitude).toBe(24);
    });

    it("should set background color when specified", async () => {
      const { create, get, batchUpdate } = getMocks();

      create.mockResolvedValue({
        data: {
          presentationId: "pres-id",
          slides: [{ objectId: "default" }],
        },
      });

      get.mockResolvedValueOnce({
        data: { slides: [makeSlide("s")] },
      });

      const config: SlideConfig = {
        title: "Test",
        slides: [
          {
            background: { red: 0.1, green: 0.2, blue: 0.3 },
            elements: [
              {
                text: "Hello",
                x: 0,
                y: 0,
                w: 100,
                h: 50,
                size: 18,
                color: "white",
              },
            ],
          },
        ],
      };

      await generateSlidesFromConfig(mockAuth, config);

      const allReqs = getAllRequests(batchUpdate);
      const bgReq = allReqs.find(
        (r) => r.updatePageProperties
      )?.updatePageProperties as Record<string, unknown>;

      expect(bgReq).toBeDefined();
      const pageProps = bgReq.pageProperties as Record<string, unknown>;
      const bgFill = pageProps.pageBackgroundFill as Record<string, unknown>;
      const solidFill = bgFill.solidFill as Record<string, unknown>;
      const color = solidFill.color as Record<string, unknown>;
      expect(color.rgbColor).toEqual({ red: 0.1, green: 0.2, blue: 0.3 });
    });

    it("should use theme colors when provided", async () => {
      const { create, get, batchUpdate } = getMocks();

      create.mockResolvedValue({
        data: {
          presentationId: "pres-id",
          slides: [{ objectId: "default" }],
        },
      });

      get.mockResolvedValueOnce({
        data: { slides: [makeSlide("s")] },
      });

      const config: SlideConfig = {
        title: "Test",
        theme: {
          colors: {
            brand: { red: 0.9, green: 0.1, blue: 0.5 },
          },
        },
        slides: [
          {
            background: "brand",
            elements: [
              {
                text: "Hello",
                x: 0,
                y: 0,
                w: 100,
                h: 50,
                size: 18,
                color: "brand",
              },
            ],
          },
        ],
      };

      await generateSlidesFromConfig(mockAuth, config);

      const allReqs = getAllRequests(batchUpdate);

      // Check background uses theme color
      const bgReq = allReqs.find(
        (r) => r.updatePageProperties
      )?.updatePageProperties as Record<string, unknown>;
      const pageProps = bgReq.pageProperties as Record<string, unknown>;
      const bgFill = pageProps.pageBackgroundFill as Record<string, unknown>;
      const solidFill = bgFill.solidFill as Record<string, unknown>;
      const bgColor = solidFill.color as Record<string, unknown>;
      expect(bgColor.rgbColor).toEqual({ red: 0.9, green: 0.1, blue: 0.5 });

      // Check text uses theme color
      const textStyle = allReqs.find(
        (r) => r.updateTextStyle
      )?.updateTextStyle as Record<string, unknown>;
      const style = textStyle.style as Record<string, unknown>;
      const fgColor = style.foregroundColor as Record<string, unknown>;
      const opaqueColor = fgColor.opaqueColor as Record<string, unknown>;
      expect(opaqueColor.rgbColor).toEqual({ red: 0.9, green: 0.1, blue: 0.5 });
    });

    it("should skip speaker notes for slides without notes", async () => {
      const { create, get, batchUpdate } = getMocks();

      create.mockResolvedValue({
        data: {
          presentationId: "pres-id",
          slides: [{ objectId: "default" }],
        },
      });

      get.mockResolvedValueOnce({
        data: {
          slides: [makeSlide("s1"), makeSlide("s2")],
        },
      });

      const config: SlideConfig = {
        title: "Test",
        slides: [
          {
            elements: [
              { text: "A", x: 0, y: 0, w: 100, h: 50, size: 18, color: "white" },
            ],
            // no notes
          },
          {
            elements: [
              { text: "B", x: 0, y: 0, w: 100, h: 50, size: 18, color: "white" },
            ],
            notes: "Only this one has notes",
          },
        ],
      };

      await generateSlidesFromConfig(mockAuth, config);

      const lastReqs = getLastRequests(batchUpdate);
      const noteInserts = lastReqs.filter(
        (r: Record<string, unknown>) =>
          r.insertText &&
          ((r.insertText as Record<string, unknown>).objectId as string).startsWith("notes_")
      );

      expect(noteInserts.length).toBe(1);
      expect(
        (noteInserts[0].insertText as Record<string, unknown>).objectId
      ).toBe("notes_s2");
    });

    it("should handle multiple elements per slide", async () => {
      const { create, get, batchUpdate } = getMocks();

      create.mockResolvedValue({
        data: {
          presentationId: "pres-id",
          slides: [{ objectId: "default" }],
        },
      });

      get.mockResolvedValueOnce({
        data: { slides: [makeSlide("s")] },
      });

      const config: SlideConfig = {
        title: "Test",
        slides: [
          {
            elements: [
              { text: "Title", x: 0, y: 0, w: 100, h: 30, size: 24, color: "white", bold: true },
              { text: "Body", x: 0, y: 40, w: 100, h: 50, size: 18, color: "white" },
              { text: "Footer", x: 0, y: 90, w: 100, h: 20, size: 12, color: "white" },
            ],
          },
        ],
      };

      await generateSlidesFromConfig(mockAuth, config);

      const allReqs = getAllRequests(batchUpdate);
      const createShapes = allReqs.filter((r) => r.createShape);
      const insertTexts = allReqs.filter(
        (r) =>
          r.insertText &&
          !((r.insertText as Record<string, unknown>).objectId as string).startsWith("notes_")
      );

      expect(createShapes.length).toBe(3);
      expect(insertTexts.length).toBe(3);
    });
  });
});

describe("update-slide mode", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("should delete existing elements on the target slide", async () => {
    const { get, batchUpdate } = getMocks();

    get.mockResolvedValueOnce({
      data: {
        slides: [
          makeSlide("slide_0"),
          makeSlide("slide_1", [
            { objectId: "old_text_1" },
            { objectId: "old_text_2" },
          ]),
        ],
      },
    });

    const config: SlideConfig = {
      title: "Q&A",
      presentationId: "pres-id",
      updateSlide: "last",
      slides: [
        {
          elements: [
            { text: "New", x: 0, y: 0, w: 100, h: 50, size: 18, color: "white" },
          ],
        },
      ],
    };

    await generateSlidesFromConfig(mockAuth, config);

    const allReqs = getAllRequests(batchUpdate);
    const deleteReqs = allReqs.filter((r) => r.deleteObject);

    expect(deleteReqs).toEqual([
      { deleteObject: { objectId: "old_text_1" } },
      { deleteObject: { objectId: "old_text_2" } },
    ]);
  });

  it("should resolve 'last' to the final slide", async () => {
    const { get, batchUpdate } = getMocks();

    get.mockResolvedValueOnce({
      data: {
        slides: [
          makeSlide("slide_0"),
          makeSlide("slide_1"),
          makeSlide("slide_2"),
        ],
      },
    });

    const config: SlideConfig = {
      title: "Q&A",
      presentationId: "pres-id",
      updateSlide: "last",
      slides: [
        {
          elements: [
            { text: "Hello", x: 0, y: 0, w: 100, h: 50, size: 18, color: "white" },
          ],
        },
      ],
    };

    await generateSlidesFromConfig(mockAuth, config);

    const allReqs = getAllRequests(batchUpdate);
    const createShape = allReqs.find((r) => r.createShape)?.createShape as Record<string, unknown>;
    const elemProps = createShape.elementProperties as Record<string, unknown>;

    // Should target slide_2 (the last slide)
    expect(elemProps.pageObjectId).toBe("slide_2");
  });

  it("should resolve numeric index to the correct slide", async () => {
    const { get, batchUpdate } = getMocks();

    get.mockResolvedValueOnce({
      data: {
        slides: [
          makeSlide("slide_0"),
          makeSlide("slide_1"),
          makeSlide("slide_2"),
        ],
      },
    });

    const config: SlideConfig = {
      title: "Q&A",
      presentationId: "pres-id",
      updateSlide: 1,
      slides: [
        {
          elements: [
            { text: "Hello", x: 0, y: 0, w: 100, h: 50, size: 18, color: "white" },
          ],
        },
      ],
    };

    await generateSlidesFromConfig(mockAuth, config);

    const allReqs = getAllRequests(batchUpdate);
    const createShape = allReqs.find((r) => r.createShape)?.createShape as Record<string, unknown>;
    const elemProps = createShape.elementProperties as Record<string, unknown>;

    expect(elemProps.pageObjectId).toBe("slide_1");
  });

  it("should create new elements from the config", async () => {
    const { get, batchUpdate } = getMocks();

    get.mockResolvedValueOnce({
      data: {
        slides: [makeSlide("target_slide")],
      },
    });

    const config: SlideConfig = {
      title: "Q&A",
      presentationId: "pres-id",
      updateSlide: 0,
      slides: [
        {
          elements: [
            { text: "Title", x: 0, y: 0, w: 100, h: 30, size: 24, color: "white", bold: true },
            { text: "Body", x: 0, y: 40, w: 100, h: 50, size: 18, color: "white" },
          ],
        },
      ],
    };

    await generateSlidesFromConfig(mockAuth, config);

    const allReqs = getAllRequests(batchUpdate);
    const createShapes = allReqs.filter((r) => r.createShape);
    const insertTexts = allReqs.filter((r) => r.insertText);

    expect(createShapes.length).toBe(2);
    expect(insertTexts.length).toBe(2);
  });

  it("should clear and replace speaker notes", async () => {
    const { get, batchUpdate } = getMocks();

    const slideWithNotes = {
      objectId: "target_slide",
      pageElements: [],
      slideProperties: {
        notesPage: {
          notesProperties: {
            speakerNotesObjectId: "notes_target_slide",
          },
          pageElements: [
            {
              objectId: "notes_target_slide",
              shape: {
                text: {
                  textElements: [
                    { textRun: { content: "Old speaker notes" } },
                  ],
                },
              },
            },
          ],
        },
      },
    };

    get.mockResolvedValueOnce({
      data: {
        slides: [slideWithNotes],
      },
    });

    const config: SlideConfig = {
      title: "Q&A",
      presentationId: "pres-id",
      updateSlide: 0,
      slides: [
        {
          elements: [
            { text: "Hello", x: 0, y: 0, w: 100, h: 50, size: 18, color: "white" },
          ],
          notes: "New speaker notes",
        },
      ],
    };

    await generateSlidesFromConfig(mockAuth, config);

    // Last batchUpdate should have deleteText + insertText for notes
    const lastReqs = getLastRequests(batchUpdate);

    const deleteText = lastReqs.find((r: Record<string, unknown>) => r.deleteText);
    expect(deleteText).toBeDefined();
    expect((deleteText.deleteText as Record<string, unknown>).objectId).toBe(
      "notes_target_slide"
    );

    const insertText = lastReqs.find((r: Record<string, unknown>) => r.insertText);
    expect(insertText).toBeDefined();
    expect((insertText.insertText as Record<string, unknown>).text).toBe(
      "New speaker notes"
    );
  });

  it("should throw if presentation has no slides", async () => {
    const { get } = getMocks();

    get.mockResolvedValueOnce({
      data: { slides: [] },
    });

    const config: SlideConfig = {
      title: "Q&A",
      presentationId: "pres-id",
      updateSlide: "last",
      slides: [
        {
          elements: [
            { text: "Hello", x: 0, y: 0, w: 100, h: 50, size: 18, color: "white" },
          ],
        },
      ],
    };

    await expect(
      generateSlidesFromConfig(mockAuth, config)
    ).rejects.toThrow("Presentation has no slides to update");
  });

  it("should throw if slide index is out of range", async () => {
    const { get } = getMocks();

    get.mockResolvedValueOnce({
      data: {
        slides: [makeSlide("only_slide")],
      },
    });

    const config: SlideConfig = {
      title: "Q&A",
      presentationId: "pres-id",
      updateSlide: 5,
      slides: [
        {
          elements: [
            { text: "Hello", x: 0, y: 0, w: 100, h: 50, size: 18, color: "white" },
          ],
        },
      ],
    };

    await expect(
      generateSlidesFromConfig(mockAuth, config)
    ).rejects.toThrow("Slide index 5 out of range");
  });

  it("should throw if config has no slides defined", async () => {
    const config: SlideConfig = {
      title: "Q&A",
      presentationId: "pres-id",
      updateSlide: "last",
      slides: [],
    };

    await expect(
      generateSlidesFromConfig(mockAuth, config)
    ).rejects.toThrow("No slides defined in config for update");
  });

  it("should update background on existing slide", async () => {
    const { get, batchUpdate } = getMocks();

    get.mockResolvedValueOnce({
      data: {
        slides: [makeSlide("target_slide")],
      },
    });

    const config: SlideConfig = {
      title: "Q&A",
      presentationId: "pres-id",
      updateSlide: 0,
      slides: [
        {
          background: { red: 0.1, green: 0.1, blue: 0.15 },
          elements: [
            { text: "Hello", x: 0, y: 0, w: 100, h: 50, size: 18, color: "white" },
          ],
        },
      ],
    };

    await generateSlidesFromConfig(mockAuth, config);

    const allReqs = getAllRequests(batchUpdate);
    const bgReq = allReqs.find(
      (r) => r.updatePageProperties
    )?.updatePageProperties as Record<string, unknown>;

    expect(bgReq).toBeDefined();
    expect(bgReq.objectId).toBe("target_slide");
  });

  it("should not create new slides (no createSlide requests)", async () => {
    const { get, batchUpdate } = getMocks();

    get.mockResolvedValueOnce({
      data: {
        slides: [makeSlide("target_slide")],
      },
    });

    const config: SlideConfig = {
      title: "Q&A",
      presentationId: "pres-id",
      updateSlide: 0,
      slides: [
        {
          elements: [
            { text: "Hello", x: 0, y: 0, w: 100, h: 50, size: 18, color: "white" },
          ],
        },
      ],
    };

    await generateSlidesFromConfig(mockAuth, config);

    const allReqs = getAllRequests(batchUpdate);
    const createSlides = allReqs.filter((r) => r.createSlide);
    expect(createSlides.length).toBe(0);
  });
});

describe("generateSlides (legacy)", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  const reviewData: ReviewData = {
    prNumber: 42,
    prTitle: "Add feature X",
    prAuthor: "testuser",
    prDate: "2024-01-15T00:00:00Z",
    repository: "org/repo",
    summary: "Added a new feature",
    changes: ["Changed file A", "Updated file B"],
    qualityAssessment: {
      codeQuality: { status: "pass", notes: "Good" },
      tests: { status: "pass", notes: "Covered" },
      security: { status: "warning", notes: "Check input" },
      performance: { status: "pass", notes: "Fast" },
    },
    issuesFound: ["Minor issue"],
    suggestions: ["Add more tests"],
    verdict: "APPROVE",
    verdictExplanation: "Looks good overall",
    businessImpact: "Improves UX",
    riskLevel: "low",
    riskFactors: ["Minor scope"],
    affectedAreas: ["Frontend"],
  };

  it("should create a presentation with 5 slides", async () => {
    const { create, get, batchUpdate } = getMocks();

    create.mockResolvedValue({
      data: {
        presentationId: "legacy-pres-id",
        slides: [{ objectId: "default_slide" }],
      },
    });

    get.mockResolvedValueOnce({
      data: {
        slides: [
          makeLegacySlide("title_1", "title_t", "title_b"),
          makeLegacySlide("summary_1", "summary_t", "summary_b"),
          makeLegacySlide("impact_1", "impact_t", "impact_b"),
          makeLegacySlide("risks_1", "risks_t", "risks_b"),
          makeLegacySlide("verdict_1", "verdict_t", "verdict_b"),
        ],
      },
    });

    const result = await generateSlides(mockAuth, reviewData);

    expect(result.presentationId).toBe("legacy-pres-id");
    expect(result.presentationUrl).toContain("legacy-pres-id");
    expect(create).toHaveBeenCalled();
  });

  it("should populate title slide content", async () => {
    const { create, get, batchUpdate } = getMocks();

    const titleTs = Date.now();
    create.mockResolvedValue({
      data: {
        presentationId: "legacy-pres-id",
        slides: [{ objectId: "default_slide" }],
      },
    });

    // The legacy function creates slides with IDs like `title_${Date.now()}`
    // We need to match those in the get response
    get.mockResolvedValueOnce({
      data: {
        slides: [],
      },
    });

    const result = await generateSlides(mockAuth, reviewData);

    // Verify batchUpdate was called with content requests
    expect(batchUpdate.mock.calls.length).toBeGreaterThanOrEqual(1);
  });

  it("should handle review data without optional fields", async () => {
    const { create, get, batchUpdate } = getMocks();

    create.mockResolvedValue({
      data: {
        presentationId: "legacy-pres-id",
        slides: [{ objectId: "default_slide" }],
      },
    });

    get.mockResolvedValueOnce({
      data: {
        slides: [
          makeLegacySlide("risks_1", "risks_t", "risks_b"),
        ],
      },
    });

    const minimalData: ReviewData = {
      prNumber: 1,
      prTitle: "Fix",
      prAuthor: "user",
      prDate: "2024-01-01",
      repository: "org/repo",
      summary: "Fix bug",
      changes: [],
      qualityAssessment: {
        codeQuality: { status: "pass", notes: "" },
        tests: { status: "pass", notes: "" },
        security: { status: "pass", notes: "" },
        performance: { status: "pass", notes: "" },
      },
      issuesFound: [],
      suggestions: [],
      verdict: "APPROVE",
      verdictExplanation: "Fine",
    };

    const result = await generateSlides(mockAuth, minimalData);
    expect(result.presentationId).toBe("legacy-pres-id");
  });
});

describe("matrix animation", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe("garbleText", () => {
    it("should preserve spaces and punctuation", () => {
      const text = "Hello, world! How are you?";
      const garbled = garbleText(text);

      expect(garbled.length).toBe(text.length);
      // Spaces preserved at same positions
      for (let i = 0; i < text.length; i++) {
        if (text[i] === " ") {
          expect(garbled[i]).toBe(" ");
        }
      }
      // Punctuation preserved
      expect(garbled[5]).toBe(",");
      expect(garbled[12]).toBe("!");
      expect(garbled[25]).toBe("?");
    });

    it("should replace non-punctuation characters with katakana", () => {
      const text = "ABC";
      const garbled = garbleText(text);

      expect(garbled).not.toBe(text);
      expect(garbled.length).toBe(3);
      // Each character should be a half-width katakana (U+FF66 to U+FF9D)
      for (const ch of garbled) {
        const code = ch.charCodeAt(0);
        expect(code).toBeGreaterThanOrEqual(0xff66);
        expect(code).toBeLessThanOrEqual(0xff9d);
      }
    });
  });

  describe("update-slide with animation", () => {
    it("should use blank text and green color for initial creation of animated elements", async () => {
      const { get, batchUpdate } = getMocks();

      get.mockResolvedValueOnce({
        data: {
          slides: [makeSlide("target_slide")],
        },
      });

      const config: SlideConfig = {
        title: "Q&A",
        presentationId: "pres-id",
        updateSlide: 0,
        slides: [
          {
            elements: [
              {
                text: "Hello world",
                x: 0, y: 0, w: 100, h: 50,
                size: 24,
                color: { red: 1, green: 1, blue: 1 },
                bold: true,
                animate: "matrix",
              },
            ],
          },
        ],
      };

      await generateSlidesFromConfig(mockAuth, config);

      // First batchUpdate = creation batch with garbled text + green color
      const creationReqs = batchUpdate.mock.calls[0][0]?.requestBody?.requests || [];
      const insertText = creationReqs.find(
        (r: Record<string, unknown>) => r.insertText
      )?.insertText as Record<string, unknown>;

      // Text should be blank (spaces for un-deposited chars)
      expect(insertText.text).not.toBe("Hello world");
      expect((insertText.text as string).length).toBe("Hello world".length);

      // Color should be Matrix green
      const textStyle = creationReqs.find(
        (r: Record<string, unknown>) => r.updateTextStyle
      )?.updateTextStyle as Record<string, unknown>;
      const style = textStyle.style as Record<string, unknown>;
      const fgColor = style.foregroundColor as Record<string, unknown>;
      const opaqueColor = fgColor.opaqueColor as Record<string, unknown>;
      expect(opaqueColor.rgbColor).toEqual({ red: 0, green: 0.8, blue: 0.2 });
    });

    it("should make animation batchUpdates based on max rain start offset", async () => {
      const { get, batchUpdate } = getMocks();

      get.mockResolvedValueOnce({
        data: {
          slides: [makeSlide("target_slide")],
        },
      });

      // "Hello" = 5 drops, offsets cycle [50,100,75] → max=100
      // lastDeposit=3, total=7 animation frames
      const config: SlideConfig = {
        title: "Q&A",
        presentationId: "pres-id",
        updateSlide: 0,
        slides: [
          {
            elements: [
              {
                text: "Hello",
                x: 0, y: 0, w: 100, h: 50,
                size: 20,
                color: { red: 1, green: 1, blue: 1 },
                animate: "matrix",
              },
            ],
          },
        ],
      };

      await generateSlidesFromConfig(mockAuth, config);

      // 1 creation batch + 10 rain animation frames = 11 total
      expect(batchUpdate.mock.calls.length).toBe(11);
    });

    it("should stagger character deposits across multiple frames", async () => {
      const { get, batchUpdate } = getMocks();

      get.mockResolvedValueOnce({
        data: {
          slides: [makeSlide("target_slide")],
        },
      });

      // "ABCD" → 4 drops, offsets [50,100,75,50]
      // Drop 0 (A): offset=50 deposits frame 1
      // Drop 1 (B): offset=100 deposits frame 3
      // Drop 2 (C): offset=75 deposits frame 2
      // Drop 3 (D): offset=50 deposits frame 1
      const config: SlideConfig = {
        title: "Q&A",
        presentationId: "pres-id",
        updateSlide: 0,
        slides: [
          {
            elements: [
              {
                text: "ABCD",
                x: 0, y: 0, w: 100, h: 50,
                size: 20,
                color: { red: 1, green: 1, blue: 1 },
                animate: "matrix",
              },
            ],
          },
        ],
      };

      await generateSlidesFromConfig(mockAuth, config);

      // Frame 1 (call index 2): A and D deposit — start black (fade progress 0)
      const frame1Reqs = batchUpdate.mock.calls[2][0]?.requestBody?.requests || [];
      const frame1MainStyles = frame1Reqs.filter(
        (r: Record<string, unknown>) =>
          r.updateTextStyle &&
          !((r.updateTextStyle as Record<string, unknown>).objectId as string).includes("_rain_")
      );
      // Base black + A(black, just deposited) + D(black, just deposited) = 3 styles
      expect(frame1MainStyles.length).toBe(3);

      // Check that A (index 0) starts black on deposit frame
      const fadeA = frame1MainStyles.find((r: Record<string, unknown>) => {
        const range = (r.updateTextStyle as Record<string, unknown>).textRange as Record<string, unknown>;
        return range.startIndex === 0 && range.endIndex === 1;
      });
      expect(fadeA).toBeDefined();
      const styleA = (fadeA!.updateTextStyle as Record<string, unknown>).style as Record<string, unknown>;
      const fgA = styleA.foregroundColor as Record<string, unknown>;
      const opA = fgA.opaqueColor as Record<string, unknown>;
      expect(opA.rgbColor).toEqual({ red: 0, green: 0, blue: 0 });

      // Frame 3 (call index 4): B deposits (black), others fading in
      const frame3Reqs = batchUpdate.mock.calls[4][0]?.requestBody?.requests || [];
      const frame3MainStyles = frame3Reqs.filter(
        (r: Record<string, unknown>) =>
          r.updateTextStyle &&
          !((r.updateTextStyle as Record<string, unknown>).objectId as string).includes("_rain_")
      );
      // Base + A(fading) + B(black) + C(fading) + D(fading) = 5 styles
      expect(frame3MainStyles.length).toBe(5);
    });

    it("should show original text fading toward final color on the last frame", async () => {
      const { get, batchUpdate } = getMocks();

      get.mockResolvedValueOnce({
        data: {
          slides: [makeSlide("target_slide")],
        },
      });

      const finalColor = { red: 1, green: 1, blue: 1 };
      const config: SlideConfig = {
        title: "Q&A",
        presentationId: "pres-id",
        updateSlide: 0,
        slides: [
          {
            elements: [
              {
                text: "Hi!",
                x: 0, y: 0, w: 100, h: 50,
                size: 20,
                color: finalColor,
                animate: "matrix",
              },
            ],
          },
        ],
      };

      await generateSlidesFromConfig(mockAuth, config);

      // Last batchUpdate call = final animation frame (per-char fade, not snap)
      const lastReqs = getLastRequests(batchUpdate);
      const insertText = lastReqs.find(
        (r: Record<string, unknown>) =>
          r.insertText &&
          !((r.insertText as Record<string, unknown>).objectId as string).includes("_rain_")
      )?.insertText as Record<string, unknown>;

      expect(insertText.text).toBe("Hi!");

      // Per-char styles: base black + 3 individual char fades = 4 styles
      const textStyles = lastReqs.filter(
        (r: Record<string, unknown>) =>
          r.updateTextStyle &&
          !((r.updateTextStyle as Record<string, unknown>).objectId as string).includes("_rain_")
      );
      expect(textStyles.length).toBe(4);

      // Each per-char style should have visible color (green > 0, faded in)
      const perCharStyles = textStyles.filter((r: Record<string, unknown>) => {
        const range = (r.updateTextStyle as Record<string, unknown>).textRange as Record<string, unknown>;
        return range.type === "FIXED_RANGE" && (range.endIndex as number) - (range.startIndex as number) === 1;
      });
      expect(perCharStyles.length).toBe(3);
      for (const pcs of perCharStyles) {
        const style = (pcs.updateTextStyle as Record<string, unknown>).style as Record<string, unknown>;
        const fgColor = style.foregroundColor as Record<string, unknown>;
        const opaqueColor = fgColor.opaqueColor as Record<string, unknown>;
        const rgb = opaqueColor.rgbColor as Record<string, number>;
        expect(rgb.green).toBeGreaterThan(0);
      }
    });

    it("should not add extra batchUpdate calls for non-animated elements", async () => {
      const { get, batchUpdate } = getMocks();

      get.mockResolvedValueOnce({
        data: {
          slides: [makeSlide("target_slide")],
        },
      });

      const config: SlideConfig = {
        title: "Q&A",
        presentationId: "pres-id",
        updateSlide: 0,
        slides: [
          {
            elements: [
              {
                text: "Static text here",
                x: 0, y: 0, w: 100, h: 50,
                size: 20,
                color: { red: 1, green: 1, blue: 1 },
                // no animate property
              },
            ],
          },
        ],
      };

      await generateSlidesFromConfig(mockAuth, config);

      // Only 1 creation batch, no animation batches
      expect(batchUpdate.mock.calls.length).toBe(1);
    });

    it("should have same frame count for single-char text", async () => {
      const { get, batchUpdate } = getMocks();

      get.mockResolvedValueOnce({
        data: {
          slides: [makeSlide("target_slide")],
        },
      });

      const finalColor = { red: 1, green: 1, blue: 1 };
      const config: SlideConfig = {
        title: "Q&A",
        presentationId: "pres-id",
        updateSlide: 0,
        slides: [
          {
            elements: [
              {
                text: "X",
                x: 0, y: 0, w: 100, h: 50,
                size: 20,
                color: finalColor,
                animate: "matrix",
              },
            ],
          },
        ],
      };

      await generateSlidesFromConfig(mockAuth, config);

      // Single char: offset=50, lastDeposit=1, totalFrames=1+1+5+1=8
      // 1 creation batch + 8 animation frames = 9 total
      expect(batchUpdate.mock.calls.length).toBe(9);

      // Final frame should have original text in final color
      const lastReqs = getLastRequests(batchUpdate);
      const insertText = lastReqs.find(
        (r: Record<string, unknown>) =>
          r.insertText &&
          !((r.insertText as Record<string, unknown>).objectId as string).includes("_rain_")
      )?.insertText as Record<string, unknown>;
      expect(insertText.text).toBe("X");
    });

    it("should create rain drops in the initial creation batch for non-space chars", async () => {
      const { get, batchUpdate } = getMocks();

      get.mockResolvedValueOnce({
        data: {
          slides: [makeSlide("target_slide")],
        },
      });

      const config: SlideConfig = {
        title: "Q&A",
        presentationId: "pres-id",
        updateSlide: 0,
        slides: [
          {
            elements: [
              {
                text: "Hello",
                x: 50, y: 100, w: 200, h: 50,
                size: 24,
                color: { red: 1, green: 1, blue: 1 },
                animate: "matrix",
              },
            ],
          },
        ],
      };

      await generateSlidesFromConfig(mockAuth, config);

      const creationReqs = batchUpdate.mock.calls[0][0]?.requestBody?.requests || [];

      // "Hello" = 5 non-space chars → 5 rain drops
      const rainCreates = creationReqs.filter(
        (r: Record<string, unknown>) =>
          r.createShape &&
          ((r.createShape as Record<string, unknown>).objectId as string).includes("_rain_")
      );
      expect(rainCreates.length).toBe(5);

      // Staggered Y: offsets cycle [50,100,75] → Y = 100 - offset
      const expectedYs = [50, 0, 25, 50, 0]; // 100-50, 100-100, 100-75, 100-50, 100-100
      for (let i = 0; i < rainCreates.length; i++) {
        const cs = rainCreates[i].createShape as Record<string, unknown>;
        const elemProps = cs.elementProperties as Record<string, unknown>;
        const transform = elemProps.transform as Record<string, unknown>;
        expect(transform.translateY).toBe(expectedYs[i] * 12700);
      }

      // First rain drop at X = pt(50 + 7.2) with TEXT_BOX_PADDING_X offset
      const rd0 = (rainCreates[0].createShape as Record<string, unknown>);
      const ep0 = rd0.elementProperties as Record<string, unknown>;
      const t0 = ep0.transform as Record<string, unknown>;
      expect(t0.translateX).toBe(57.2 * 12700);

      // Rain drop insertTexts should have katakana characters
      const rainInserts = creationReqs.filter(
        (r: Record<string, unknown>) =>
          r.insertText &&
          ((r.insertText as Record<string, unknown>).objectId as string).includes("_rain_")
      );
      expect(rainInserts.length).toBe(5);
      for (const ri of rainInserts) {
        const text = (ri.insertText as Record<string, unknown>).text as string;
        expect(text.length).toBe(1);
        const code = text.charCodeAt(0);
        expect(code).toBeGreaterThanOrEqual(0xff66);
        expect(code).toBeLessThanOrEqual(0xff9d);
      }

      // All rain drops start black (they fade in during animation)
      const rainStyles = creationReqs.filter(
        (r: Record<string, unknown>) =>
          r.updateTextStyle &&
          ((r.updateTextStyle as Record<string, unknown>).objectId as string).includes("_rain_")
      );
      expect(rainStyles.length).toBe(5);
      for (const rs of rainStyles) {
        const style = (rs.updateTextStyle as Record<string, unknown>).style as Record<string, unknown>;
        const fgColor = style.foregroundColor as Record<string, unknown>;
        const opaqueColor = fgColor.opaqueColor as Record<string, unknown>;
        expect(opaqueColor.rgbColor).toEqual({ red: 0, green: 0, blue: 0 });
      }
    });

    it("should move all rain drops simultaneously via updatePageElementTransform", async () => {
      const { get, batchUpdate } = getMocks();

      get.mockResolvedValueOnce({
        data: {
          slides: [makeSlide("target_slide")],
        },
      });

      const config: SlideConfig = {
        title: "Q&A",
        presentationId: "pres-id",
        updateSlide: 0,
        slides: [
          {
            elements: [
              {
                text: "ABC",
                x: 10, y: 80, w: 200, h: 50,
                size: 20,
                color: { red: 1, green: 1, blue: 1 },
                animate: "matrix",
              },
            ],
          },
        ],
      };

      await generateSlidesFromConfig(mockAuth, config);

      // Frame 0 (call index 1): all 3 rain drops should move simultaneously
      const frame0Reqs = batchUpdate.mock.calls[1][0]?.requestBody?.requests || [];
      const transformReqs = frame0Reqs.filter(
        (r: Record<string, unknown>) => r.updatePageElementTransform
      );

      // All 3 rain drops have transforms
      expect(transformReqs.length).toBe(3);

      // Staggered Y: rainY = elemY - offset + (frame+1)*25
      // Frame 0: [80-50+25=55, 80-100+25=5, 80-75+25=30]
      const expectedYs = [55, 5, 30];
      for (let i = 0; i < transformReqs.length; i++) {
        const t = transformReqs[i].updatePageElementTransform as Record<string, unknown>;
        expect(t.applyMode).toBe("ABSOLUTE");
        const transform = t.transform as Record<string, unknown>;
        expect(transform.translateY).toBe(expectedYs[i] * 12700);
      }

      // X positions differ per character (with TEXT_BOX_PADDING_X = 7.2)
      const charWidth = 20 * 0.48; // = 9.6
      const padX = 7.2;
      const xValues = transformReqs.map((tr: Record<string, unknown>) => {
        const t = tr.updatePageElementTransform as Record<string, unknown>;
        const transform = t.transform as Record<string, unknown>;
        return transform.translateX;
      });
      expect(xValues[0]).toBe((10 + padX + 0 * charWidth) * 12700);
      expect(xValues[1]).toBe((10 + padX + 1 * charWidth) * 12700);
      expect(xValues[2]).toBe((10 + padX + 2 * charWidth) * 12700);
    });

    it("should show katakana in rain drops on pre-deposit frames", async () => {
      const { get, batchUpdate } = getMocks();

      get.mockResolvedValueOnce({
        data: {
          slides: [makeSlide("target_slide")],
        },
      });

      const config: SlideConfig = {
        title: "Q&A",
        presentationId: "pres-id",
        updateSlide: 0,
        slides: [
          {
            elements: [
              {
                text: "ABCD",
                x: 0, y: 50, w: 200, h: 50,
                size: 20,
                color: { red: 1, green: 1, blue: 1 },
                animate: "matrix",
              },
            ],
          },
        ],
      };

      await generateSlidesFromConfig(mockAuth, config);

      // Frame 0 (call index 1): all 4 rain drops should have katakana text
      const frame0Reqs = batchUpdate.mock.calls[1][0]?.requestBody?.requests || [];
      const rainInserts = frame0Reqs.filter(
        (r: Record<string, unknown>) =>
          r.insertText &&
          ((r.insertText as Record<string, unknown>).objectId as string).includes("_rain_")
      );
      expect(rainInserts.length).toBe(4);
      for (const ri of rainInserts) {
        const text = (ri.insertText as Record<string, unknown>).text as string;
        expect(text.length).toBe(1);
        const code = text.charCodeAt(0);
        expect(code).toBeGreaterThanOrEqual(0xff66);
        expect(code).toBeLessThanOrEqual(0xff9d);
      }
    });

    it("should delete all rain drops on the final animation frame", async () => {
      const { get, batchUpdate } = getMocks();

      get.mockResolvedValueOnce({
        data: {
          slides: [makeSlide("target_slide")],
        },
      });

      const config: SlideConfig = {
        title: "Q&A",
        presentationId: "pres-id",
        updateSlide: 0,
        slides: [
          {
            elements: [
              {
                text: "AB",
                x: 0, y: 50, w: 200, h: 50,
                size: 20,
                color: { red: 1, green: 1, blue: 1 },
                animate: "matrix",
              },
            ],
          },
        ],
      };

      await generateSlidesFromConfig(mockAuth, config);

      // "AB" = 2 drops, offsets [50,100], max=100 → 10 anim frames, 11 total
      expect(batchUpdate.mock.calls.length).toBe(11);

      // Last frame should have deleteObject for all 2 rain drops
      const lastReqs = getLastRequests(batchUpdate);
      const deleteObjs = lastReqs.filter(
        (r: Record<string, unknown>) =>
          r.deleteObject &&
          ((r.deleteObject as Record<string, unknown>).objectId as string).includes("_rain_")
      );
      expect(deleteObjs.length).toBe(2);
    });

    it("should not create rain drops for space characters", async () => {
      const { get, batchUpdate } = getMocks();

      get.mockResolvedValueOnce({
        data: {
          slides: [makeSlide("target_slide")],
        },
      });

      const config: SlideConfig = {
        title: "Q&A",
        presentationId: "pres-id",
        updateSlide: 0,
        slides: [
          {
            elements: [
              {
                text: "A B",
                x: 0, y: 50, w: 200, h: 50,
                size: 20,
                color: { red: 1, green: 1, blue: 1 },
                animate: "matrix",
              },
            ],
          },
        ],
      };

      await generateSlidesFromConfig(mockAuth, config);

      const creationReqs = batchUpdate.mock.calls[0][0]?.requestBody?.requests || [];
      // "A B" = 2 non-space chars → 2 rain drops (space skipped)
      const rainCreates = creationReqs.filter(
        (r: Record<string, unknown>) =>
          r.createShape &&
          ((r.createShape as Record<string, unknown>).objectId as string).includes("_rain_")
      );
      expect(rainCreates.length).toBe(2);
    });

    it("should fade rain drops at different rates after deposit", async () => {
      const { get, batchUpdate } = getMocks();

      get.mockResolvedValueOnce({
        data: {
          slides: [makeSlide("target_slide")],
        },
      });

      const config: SlideConfig = {
        title: "Q&A",
        presentationId: "pres-id",
        updateSlide: 0,
        slides: [
          {
            elements: [
              {
                text: "AB",
                x: 0, y: 50, w: 200, h: 50,
                size: 20,
                color: { red: 1, green: 1, blue: 1 },
                animate: "matrix",
              },
            ],
          },
        ],
      };

      await generateSlidesFromConfig(mockAuth, config);

      // Frame 5 (call index 6): both drops deposited, fading at different rates
      // Drop 0 (offset=50, fadeSteps=3): 4 frames since deposit → fully black
      // Drop 1 (offset=100, fadeSteps=5): 2 frames since deposit → partially faded
      const frame5Reqs = batchUpdate.mock.calls[6][0]?.requestBody?.requests || [];
      const rainStyles = frame5Reqs.filter(
        (r: Record<string, unknown>) =>
          r.updateTextStyle &&
          ((r.updateTextStyle as Record<string, unknown>).objectId as string).includes("_rain_")
      );
      expect(rainStyles.length).toBe(2);

      // Extract green channel from each rain drop
      const greens = rainStyles.map((rs: Record<string, unknown>) => {
        const style = (rs.updateTextStyle as Record<string, unknown>).style as Record<string, unknown>;
        const fgColor = style.foregroundColor as Record<string, unknown>;
        const opaqueColor = fgColor.opaqueColor as Record<string, unknown>;
        const rgb = opaqueColor.rgbColor as Record<string, number>;
        return rgb.green;
      });

      // Drop 0: fully faded to black
      expect(greens[0]).toBe(0);

      // Drop 1: partially faded (less than matrix green but not black)
      expect(greens[1]).toBeGreaterThan(0);
      expect(greens[1]).toBeLessThan(0.8);

      // Different fade rates produce different values
      expect(greens[0]).not.toBe(greens[1]);
    });
  });
});
