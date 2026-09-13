import type { Issue, Party } from "../domain/models";

export interface IssueSheetRenderInput {
  issue: Issue;
  parties: Party[];
}

export async function renderIssueSheet(_input: IssueSheetRenderInput): Promise<void> {
  throw new Error("Not implemented");
}
