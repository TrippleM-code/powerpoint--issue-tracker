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

  initializeIssuePanel();
});
