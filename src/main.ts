import "./styles.css";
import { initializeIssuePanel } from "./ui/issue-panel";

declare const Office: any;

Office.onReady((info: any) => {
  if (info.host !== Office.HostType.PowerPoint) {
    const banner = document.getElementById("statusBanner");
    if (banner) {
      banner.textContent = "IssueFlow must be opened in Microsoft PowerPoint.";
      banner.className = "banner error";
    }
    return;
  }

  initializeIssuePanel().catch((error: unknown) => {
    const banner = document.getElementById("statusBanner");
    if (banner) {
      banner.textContent = error instanceof Error ? error.message : String(error);
      banner.className = "banner error";
    }
    console.error("IssueFlow initialization failed:", error);
  });
});
