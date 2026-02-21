import { codeToHtml } from "shiki";
import { CopyButton } from "./copy-button";

const SHELL_LANGS = new Set(["bash", "sh", "shell", "zsh"]);

/**
 * Clean shell code for clipboard — strip comments, prompts, and blank noise
 * so users get runnable commands when they paste into a terminal.
 * Visual display is unaffected (comments still render in the code block).
 */
function cleanShellForCopy(raw: string): string {
  return raw
    .split("\n")
    .filter((line) => !/^\s*#/.test(line))             // drop comment-only lines
    .map((line) => line.replace(/\s{2,}#\s.*$/, ""))   // strip inline comments (padded `  # ...`)
    .map((line) => line.replace(/^\s*\$\s/, ""))        // strip `$ ` prompt prefix
    .map((line) => line.trimEnd())                      // trim trailing whitespace
    .filter((line, i, arr) =>                           // collapse consecutive blank lines
      !(line === "" && (i === 0 || arr[i - 1] === ""))
    )
    .join("\n")
    .trim();
}

interface CodeBlockProps {
  code: string;
  lang?: string;
}

export async function CodeBlock({ code, lang = "bash" }: CodeBlockProps) {
  const trimmed = code.trim();
  const copyText = SHELL_LANGS.has(lang) ? cleanShellForCopy(trimmed) : trimmed;

  let html = await codeToHtml(trimmed, {
    lang,
    themes: {
      light: "github-light-default",
      dark: "github-dark-default",
    },
    defaultColor: false,
    colorReplacements: {
      "github-dark-default": { "#ffa657": "#4ade80" },
      "github-light-default": { "#953800": "#16a34a" },
    },
  });

  // Colour CLI command names emerald green
  const cliNames = [
    "agent-xlsx",
    "uvx", "uv", "pip", "pipx", "npx",
    "probe", "read", "search", "export", "overview", "inspect",
    "format", "write", "sheet", "screenshot", "objects", "recalc",
    "vba", "license",
  ];
  const cliPattern = new RegExp(
    `(<span style="[^"]*">)(${cliNames.join("|")})(</span>)`,
    "g"
  );
  html = html.replace(cliPattern, '<span style="color:#16a34a;--shiki-dark:#4ade80">$2</span>');

  return (
    <div className="code-block relative group">
      <CopyButton code={copyText} />
      <div dangerouslySetInnerHTML={{ __html: html }} />
    </div>
  );
}
