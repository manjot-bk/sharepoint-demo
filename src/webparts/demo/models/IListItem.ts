export interface IListItem {
  Id: number;
  Title: string;
  Description?: string;
  Created: string;
  Modified: string;
  Author: {
    Title: string;
    Email: string;
  };
}

export interface ITaskItem extends IListItem {
  Status: string;
  Priority: string;
  DueDate?: string;
  AssignedTo?: {
    Title: string;
    Email: string;
  };
}

export interface IDocumentItem extends IListItem {
  FileRef: string;
  FileSize: number;
  FileType: string;
}
