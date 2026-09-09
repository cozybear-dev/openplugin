import type { HostKind } from "@openplugin/core";
import { Children, type ReactNode } from "react";
import ReactMarkdown, { defaultUrlTransform } from "react-markdown";
import remarkGfm from "remark-gfm";
import { jumpTo, splitCitations } from "../citations";

export function MarkdownMessage(props: { hostKind: HostKind; text: string }) {
  return (
    <div className="op-md">
      <ReactMarkdown
        remarkPlugins={[remarkGfm]}
        urlTransform={safeUrl}
        components={{
          a: ({ href, children }) => (
            <a href={href} target="_blank" rel="noreferrer">
              {children}
            </a>
          ),
          img: () => null,
          p: ({ children }) => <p>{cite(props.hostKind, children)}</p>,
          li: ({ children }) => <li>{cite(props.hostKind, children)}</li>,
          td: ({ children }) => <td>{cite(props.hostKind, children)}</td>,
          th: ({ children }) => <th>{cite(props.hostKind, children)}</th>,
          h1: ({ children }) => <h1>{cite(props.hostKind, children)}</h1>,
          h2: ({ children }) => <h2>{cite(props.hostKind, children)}</h2>,
          h3: ({ children }) => <h3>{cite(props.hostKind, children)}</h3>,
          h4: ({ children }) => <h4>{cite(props.hostKind, children)}</h4>,
          blockquote: ({ children }) => <blockquote>{cite(props.hostKind, children)}</blockquote>,
          strong: ({ children }) => <strong>{cite(props.hostKind, children)}</strong>,
          em: ({ children }) => <em>{cite(props.hostKind, children)}</em>,
          del: ({ children }) => <del>{cite(props.hostKind, children)}</del>
        }}
      >
        {props.text}
      </ReactMarkdown>
    </div>
  );
}

function safeUrl(url: string): string {
  const next = defaultUrlTransform(url);
  if (!next) return "";
  try {
    const parsed = new URL(next, "https://invalid.local");
    if (parsed.protocol === "http:" || parsed.protocol === "https:" || parsed.protocol === "mailto:") return next;
  } catch {
    return "";
  }
  return "";
}

function cite(hostKind: HostKind, children: ReactNode): ReactNode {
  return Children.map(children, (child) => {
    if (typeof child !== "string") return child;
    const parts = splitCitations(child);
    if (parts.length === 1 && !parts[0].citation) return child;
    return parts.map((part, i) =>
      part.citation ? (
        <button
          key={i}
          type="button"
          className="op-cite"
          onClick={() => void jumpTo(hostKind, part.citation!)}
        >
          {part.text}
        </button>
      ) : (
        <span key={i}>{part.text}</span>
      )
    );
  });
}
