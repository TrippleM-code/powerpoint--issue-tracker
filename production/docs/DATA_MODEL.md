# Data Model

## Schema version

`issueflow.schemaVersion = 1`

## Issue

```ts
interface Issue {
  id: string;
  areaCode: string;
  roomSpace: string;
  description: string;
  createdAt: string;
  updatedAt?: string;
  actions: Action[];
}
```

## Action

```ts
interface Action {
  id: string;
  party: string;
  required: string;
  status: string;
  createdAt: string;
  updatedAt?: string;
}
```

## Presentation settings

```ts
interface PresentationSettings {
  schemaVersion: number;
  parties: Party[];
  statuses: StatusDefinition[];
}
```

## Party

```ts
interface Party {
  id: string;
  name: string;
  logoDataUrl?: string;
}
```

## Source of truth

Structured metadata is authoritative.
Generated PowerPoint shapes are a visual projection of that data.
Manual Reference Images content is user-owned and must never be deleted by refresh.
