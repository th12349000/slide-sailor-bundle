import { describe, expect, it } from "vitest";
import { localAiErrorMessage } from "@/lib/local-ai";

describe("local AI errors", () => {
  it("explains Safari's generic load failure as a first-use model download issue", () => {
    expect(localAiErrorMessage(new TypeError("Load failed"))).toContain(
      "Connect once to the internet",
    );
  });

  it("preserves useful details for device runtime failures", () => {
    expect(localAiErrorMessage(new Error("WebGPU out of memory"))).toContain(
      "WebGPU out of memory",
    );
  });
});
