export type SessionResponse = {
  rootPath: string | null;
};

export type PickRootResponse = {
  rootPath: string | null;
};

export type ErrorResponse = {
  error: string;
};

export type BrowseEntry = {
  name: string;
  relativePath: string;
  isDirectory: boolean;
  sizeBytes: number | null;
  modifiedUtc: string;
};

export type BrowseResponse = {
  entries: BrowseEntry[];
};

export type FileKind =
  | "pdf"
  | "image"
  | "html"
  | "json"
  | "text"
  | "other";

export type FileMetaResponse = {
  relativePath: string;
  name: string;
  kind: FileKind;
  sizeBytes: number;
  modifiedUtc: string;
  mimeType: string;
};
