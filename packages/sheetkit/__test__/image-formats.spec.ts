import { existsSync, unlinkSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { Workbook } from "../index.js";

const TEST_DIR = import.meta.dirname;

function tmpFile(name: string) {
  return join(TEST_DIR, name);
}

function cleanup(...files: string[]) {
  for (const f of files) {
    if (existsSync(f)) unlinkSync(f);
  }
}

// Minimal 1x1 PNG for testing image insertion
const DUMMY_IMAGE = Buffer.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d,
  0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
  0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4, 0x89, 0x00, 0x00, 0x00,
  0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49,
  0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
]);

function addTestImage(wb: InstanceType<typeof Workbook>, format: string) {
  wb.addImage("Sheet1", {
    data: DUMMY_IMAGE,
    format,
    fromCell: "A1",
    widthPx: 100,
    heightPx: 100,
  });
}

// existing formats regression
describe("Image formats - existing formats", () => {
  const out = tmpFile("test-img-existing.xlsx");
  afterEach(() => cleanup(out));

  it("should accept png format", async () => {
    const wb = new Workbook();
    addTestImage(wb, "png");
    await wb.save(out);
    expect(existsSync(out)).toBe(true);
  });

  it("should accept jpeg format", async () => {
    const wb = new Workbook();
    addTestImage(wb, "jpeg");
    await wb.save(out);
    expect(existsSync(out)).toBe(true);
  });

  it("should accept jpg alias", async () => {
    const wb = new Workbook();
    addTestImage(wb, "jpg");
    await wb.save(out);
    expect(existsSync(out)).toBe(true);
  });

  it("should accept gif format", async () => {
    const wb = new Workbook();
    addTestImage(wb, "gif");
    await wb.save(out);
    expect(existsSync(out)).toBe(true);
  });
});

// new formats
describe("Image formats - new formats", () => {
  const out = tmpFile("test-img-new.xlsx");
  afterEach(() => cleanup(out));

  it("should accept bmp format", async () => {
    const wb = new Workbook();
    addTestImage(wb, "bmp");
    await wb.save(out);
    expect(existsSync(out)).toBe(true);
  });

  it("should accept ico format", async () => {
    const wb = new Workbook();
    addTestImage(wb, "ico");
    await wb.save(out);
    expect(existsSync(out)).toBe(true);
  });

  it("should accept tiff format", async () => {
    const wb = new Workbook();
    addTestImage(wb, "tiff");
    await wb.save(out);
    expect(existsSync(out)).toBe(true);
  });

  it("should accept tif alias", async () => {
    const wb = new Workbook();
    addTestImage(wb, "tif");
    await wb.save(out);
    expect(existsSync(out)).toBe(true);
  });

  it("should accept svg format", async () => {
    const wb = new Workbook();
    addTestImage(wb, "svg");
    await wb.save(out);
    expect(existsSync(out)).toBe(true);
  });

  it("should accept emf format", async () => {
    const wb = new Workbook();
    addTestImage(wb, "emf");
    await wb.save(out);
    expect(existsSync(out)).toBe(true);
  });

  it("should accept emz format", async () => {
    const wb = new Workbook();
    addTestImage(wb, "emz");
    await wb.save(out);
    expect(existsSync(out)).toBe(true);
  });

  it("should accept wmf format", async () => {
    const wb = new Workbook();
    addTestImage(wb, "wmf");
    await wb.save(out);
    expect(existsSync(out)).toBe(true);
  });

  it("should accept wmz format", async () => {
    const wb = new Workbook();
    addTestImage(wb, "wmz");
    await wb.save(out);
    expect(existsSync(out)).toBe(true);
  });
});

// case-insensitive format strings
describe("Image formats - case insensitive", () => {
  const out = tmpFile("test-img-case.xlsx");
  afterEach(() => cleanup(out));

  it("should accept uppercase PNG", async () => {
    const wb = new Workbook();
    addTestImage(wb, "PNG");
    await wb.save(out);
    expect(existsSync(out)).toBe(true);
  });

  it("should accept mixed case Svg", async () => {
    const wb = new Workbook();
    addTestImage(wb, "Svg");
    await wb.save(out);
    expect(existsSync(out)).toBe(true);
  });

  it("should accept uppercase TIFF", async () => {
    const wb = new Workbook();
    addTestImage(wb, "TIFF");
    await wb.save(out);
    expect(existsSync(out)).toBe(true);
  });

  it("should accept uppercase JPEG", async () => {
    const wb = new Workbook();
    addTestImage(wb, "JPEG");
    await wb.save(out);
    expect(existsSync(out)).toBe(true);
  });

  it("should accept mixed case Emf", async () => {
    const wb = new Workbook();
    addTestImage(wb, "Emf");
    await wb.save(out);
    expect(existsSync(out)).toBe(true);
  });
});

// unsupported format error
describe("Image formats - unsupported", () => {
  it("should reject unsupported format string", () => {
    const wb = new Workbook();
    expect(() => addTestImage(wb, "webp")).toThrow();
  });

  it("should reject empty format string", () => {
    const wb = new Workbook();
    expect(() => addTestImage(wb, "")).toThrow();
  });

  it("should reject arbitrary string", () => {
    const wb = new Workbook();
    expect(() => addTestImage(wb, "notaformat")).toThrow();
  });
});
