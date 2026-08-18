import { spawn } from "node:child_process";
import { join } from "node:path";
import process from "node:process";
import { parseWorkIqCliAnswer, WorkIqCliAnswer } from "./workiq";

const WORK_IQ_TIMEOUT_MS = 5 * 60 * 1000;
const WORK_IQ_MAX_OUTPUT_BYTES = 10 * 1024 * 1024;

interface WorkIqRuntime {
  platform: NodeJS.Platform;
  arch: string;
  appData?: string;
}

export function resolveWorkIqExecutable(configuredPath: string, runtime: WorkIqRuntime = getRuntime()): string {
  const trimmedPath = configuredPath.trim();

  if (trimmedPath) {
    return trimmedPath;
  }

  if (runtime.platform !== "win32") {
    return "workiq";
  }

  if (!runtime.appData) {
    throw new Error("APPDATA is unavailable. Set the Work IQ executable path in the plugin settings.");
  }

  const platformDirectory = runtime.arch === "arm64" ? "win-arm64" : "win-x64";
  return join(
    runtime.appData,
    "npm",
    "node_modules",
    "@microsoft",
    "workiq",
    "bin",
    platformDirectory,
    "workiq.exe"
  );
}

export async function askWorkIq(prompt: string, configuredPath: string): Promise<WorkIqCliAnswer> {
  const question = prompt.trim();

  if (!question) {
    throw new Error("Enter a Work IQ prompt.");
  }

  const executable = resolveWorkIqExecutable(configuredPath);
  const output = await executeWorkIq(executable, ["ask", "--json", "--question", question]);

  try {
    return parseWorkIqCliAnswer(JSON.parse(output));
  } catch (error) {
    if (error instanceof SyntaxError) {
      throw new Error("Work IQ CLI returned invalid JSON.");
    }

    throw error;
  }
}

export function executeWorkIq(
  executable: string,
  args: string[],
  timeoutMs = WORK_IQ_TIMEOUT_MS
): Promise<string> {
  return new Promise((resolve, reject) => {
    const child = spawn(executable, args, {
      stdio: ["ignore", "pipe", "pipe"],
      windowsHide: true
    });
    let stdout = "";
    let stderr = "";
    let settled = false;

    const finish = (action: () => void): void => {
      if (settled) {
        return;
      }

      settled = true;
      clearTimeout(timeout);
      child.kill();
      action();
    };

    const timeout = setTimeout(() => {
      const detail = stderr.trim() || stdout.trim();
      finish(() => reject(new Error(`Work IQ CLI timed out.${detail ? ` ${detail}` : ""}`)));
    }, timeoutMs);

    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");

    child.stdout.on("data", (chunk: string) => {
      stdout += chunk;

      if (Buffer.byteLength(stdout, "utf8") > WORK_IQ_MAX_OUTPUT_BYTES) {
        finish(() => reject(new Error("Work IQ CLI returned too much output.")));
        return;
      }

      try {
        JSON.parse(stdout);
        finish(() => resolve(stdout.trim()));
      } catch {
        // JSON can span multiple stdout chunks; keep reading until it is complete.
      }
    });

    child.stderr.on("data", (chunk: string) => {
      stderr += chunk;
    });

    child.on("error", (error) => {
      finish(() => reject(new Error(`Could not run Work IQ CLI. ${error.message}`)));
    });

    child.on("exit", (code) => {
      if (settled) {
        return;
      }

      const detail = stderr.trim() || stdout.trim() || `Process exited with code ${code ?? "unknown"}.`;
      finish(() => reject(new Error(`Could not run Work IQ CLI. ${detail}`)));
    });
  });
}

function getRuntime(): WorkIqRuntime {
  return {
    platform: process.platform,
    arch: process.arch,
    appData: process.env.APPDATA
  };
}