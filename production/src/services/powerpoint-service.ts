// Production implementation will wrap Office.js calls here.
// UI and domain modules must not call PowerPoint.run directly.

export class PowerPointService {
  async getSelectedSlideId(): Promise<string> {
    throw new Error("Not implemented");
  }

  async selectSlideById(_slideId: string): Promise<void> {
    throw new Error("Not implemented");
  }
}
