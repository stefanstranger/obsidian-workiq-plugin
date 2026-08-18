import { describe, expect, it } from "vitest";
import {
  formatWorkIqCliAnswer,
  parseWorkIqCliAnswer,
} from "../src/workiq";
import process from "node:process";
import { executeWorkIq, resolveWorkIqExecutable } from "../src/workiq-cli";

describe("Work IQ CLI helpers", () => {
  it("parses the official CLI JSON response", () => {
    expect(
      parseWorkIqCliAnswer({
        isError: false,
        response: "JSON integration successful.",
        conversationId: "conversation-123",
        agentId: "bizchat-as-gpt-scenario"
      })
    ).toEqual({
      conversationId: "conversation-123",
      response: "JSON integration successful."
    });
  });

  it("formats the CLI answer as Markdown", () => {
    expect(
      formatWorkIqCliAnswer("Summarize project alpha", {
        conversationId: "conversation-123",
        response: "Project Alpha is on track."
      })
    ).toBe("## Work IQ: Summarize project alpha\n\nProject Alpha is on track.");
  });

  it("surfaces CLI error responses", () => {
    expect(() =>
      parseWorkIqCliAnswer({
        isError: true,
        response: "Sign-in failed.",
        conversationId: ""
      })
    ).toThrow("Sign-in failed.");
  });

  it("resolves the signed Windows executable from the global npm location", () => {
    expect(
      resolveWorkIqExecutable("", {
        platform: "win32",
        arch: "x64",
        appData: "C:\\Users\\stefstr\\AppData\\Roaming"
      })
    ).toBe(
      "C:\\Users\\stefstr\\AppData\\Roaming\\npm\\node_modules\\@microsoft\\workiq\\bin\\win-x64\\workiq.exe"
    );
  });

  it("uses an explicitly configured executable path", () => {
    expect(
      resolveWorkIqExecutable(" C:\\Tools\\workiq.exe ", {
        platform: "win32",
        arch: "x64"
      })
    ).toBe("C:\\Tools\\workiq.exe");
  });

  it("returns complete JSON without waiting for the CLI process to exit", async () => {
    const output = await executeWorkIq(
      process.execPath,
      [
        "-e",
        "process.stdout.write(JSON.stringify({ isError: false, response: 'ready', conversationId: '123' })); setInterval(() => {}, 1000);"
      ],
      1000
    );

    expect(JSON.parse(output)).toEqual({
      isError: false,
      response: "ready",
      conversationId: "123"
    });
  });
});
