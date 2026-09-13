export interface ActionItem {
  id: string;
  party: string;
  required: string;
  status: string;
  createdAt: string;
  updatedAt?: string;
}

export interface Issue {
  id: string;
  areaCode: string;
  roomSpace: string;
  description: string;
  createdAt: string;
  updatedAt?: string;
  actions: ActionItem[];
}
