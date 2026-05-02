export type SessionResponse = {
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
