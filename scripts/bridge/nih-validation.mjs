import path from "node:path";
import { mkdir, writeFile } from "node:fs/promises";

const ROOT_DIR = path.resolve(new URL("../..", import.meta.url).pathname);
const DOCS_DIR = path.join(ROOT_DIR, "docs");
const ARTIFACTS_DIR = path.join(ROOT_DIR, "scripts", "bridge", "artifacts");

export const NIH_DOC_LANES = [
  {
    lane: "copy-a",
    aliases: ["a", "copy-a", "copya", "medium"],
    fileName: "NIH_Biocoating_CraftV6.docx",
    readOnly: false,
    role: "Copy A / medium mutation scenarios",
  },
  {
    lane: "copy-b",
    aliases: ["b", "copy-b", "copyb", "hard"],
    fileName: "NIH_Biocoating_CraftV6 copy.docx",
    readOnly: false,
    role: "Copy B / hard mutation + review scenarios",
  },
  {
    lane: "source-of-truth",
    aliases: ["source", "source-of-truth", "reference", "readonly"],
    fileName: "NIH_Biocoating_CraftV6 source-of-truth.docx",
    readOnly: true,
    role: "Untouched reference",
  },
];

export function parseValidationLane(input) {
  const normalized = String(input ?? "copy-a").trim().toLowerCase();
  const lane = NIH_DOC_LANES.find((candidate) =>
    candidate.aliases.includes(normalized),
  );
  if (!lane) {
    throw new Error(
      `Unknown NIH validation lane \"${input}\". Supported lanes: ${NIH_DOC_LANES.map((candidate) => candidate.lane).join(", ")}.`,
    );
  }
  return lane;
}

export function buildEvidencePaths({ rootDir = ROOT_DIR, lane, scenario, timestamp }) {
  const safeScenario = String(scenario || "manual-check")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, "-")
    .replace(/^-+|-+$/g, "") || "manual-check";
  const safeTimestamp = String(timestamp || new Date().toISOString())
    .replace(/[:.]/g, "-")
    .replace(/[^a-zA-Z0-9TZ_-]/g, "");
  const baseDir = path.join(
    rootDir,
    "scripts",
    "bridge",
    "artifacts",
    lane.lane,
    safeScenario,
    safeTimestamp,
  );
  return {
    baseDir,
    preReadbackPath: path.join(baseDir, "pre-readback.json"),
    postReadbackPath: path.join(baseDir, "post-readback.json"),
    liveSessionPath: path.join(baseDir, "live-session.json"),
    offlineInspectPath: path.join(baseDir, "offline-inspect.txt"),
    summaryPath: path.join(baseDir, "summary.json"),
  };
}

function parseArgs(argv) {
  const options = {
    lane: "copy-a",
    scenario: "manual-check",
    timestamp: undefined,
    printJson: false,
    writeSummary: true,
  };

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === "--lane") options.lane = argv[++index];
    else if (arg === "--scenario") options.scenario = argv[++index];
    else if (arg === "--timestamp") options.timestamp = argv[++index];
    else if (arg === "--print-json") options.printJson = true;
    else if (arg === "--no-summary") options.writeSummary = false;
    else if (arg === "--help" || arg === "-h") options.help = true;
    else throw new Error(`Unknown argument: ${arg}`);
  }

  return options;
}

function buildCommandHints({ lane, evidence }) {
  return {
    liveBridgeReadback:
      `pnpm exec office-bridge inspect word --compact > \"${evidence.liveSessionPath}\"`,
    offlineInspection:
      `python scripts/officecli_inspect.py \"${path.join(DOCS_DIR, lane.fileName)}\" > \"${evidence.offlineInspectPath}\"`,
    preReadContract:
      `Save exact pre-read scope JSON to \"${evidence.preReadbackPath}\" before mutation.`,
    postReadContract:
      `Save exact post-read scope JSON to \"${evidence.postReadbackPath}\" after mutation and compare literals before declaring success.`,
  };
}

export async function prepareValidationLane(options = {}) {
  const lane = parseValidationLane(options.lane);
  const evidence = buildEvidencePaths({
    rootDir: options.rootDir || ROOT_DIR,
    lane,
    scenario: options.scenario,
    timestamp: options.timestamp,
  });
  await mkdir(evidence.baseDir, { recursive: true });

  const summary = {
    lane: lane.lane,
    role: lane.role,
    readOnly: lane.readOnly,
    editedFile: path.join(DOCS_DIR, lane.fileName),
    sourceOfTruthFile: path.join(
      DOCS_DIR,
      "NIH_Biocoating_CraftV6 source-of-truth.docx",
    ),
    artifactDir: evidence.baseDir,
    evidence,
    commands: buildCommandHints({ lane, evidence }),
  };

  if (options.writeSummary !== false) {
    await writeFile(evidence.summaryPath, `${JSON.stringify(summary, null, 2)}\n`, "utf8");
  }

  return summary;
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  if (options.help) {
    process.stdout.write(`Usage: node scripts/bridge/nih-validation.mjs [--lane copy-a|copy-b|source-of-truth] [--scenario name] [--timestamp iso] [--print-json]\n`);
    return;
  }
  const summary = await prepareValidationLane(options);
  if (options.printJson) {
    process.stdout.write(`${JSON.stringify(summary, null, 2)}\n`);
    return;
  }
  process.stdout.write(`Prepared ${summary.lane} lane for ${summary.editedFile}\n`);
  process.stdout.write(`Artifacts: ${summary.artifactDir}\n`);
  process.stdout.write(`Summary: ${summary.evidence.summaryPath}\n`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
    process.exit(1);
  });
}
