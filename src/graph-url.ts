export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

export function continuationPath(next: string, initial: string, errorMessage: string): string {
  let url: URL;
  try {
    url = new URL(next);
  } catch {
    throw new Error(errorMessage);
  }
  const base = new URL(`${GRAPH_BASE}${initial}`);
  if (url.origin !== base.origin || url.pathname !== base.pathname || url.username || url.password || url.hash || url.href.endsWith("#")) {
    throw new Error(errorMessage);
  }
  return url.pathname.slice("/v1.0".length) + url.search;
}
