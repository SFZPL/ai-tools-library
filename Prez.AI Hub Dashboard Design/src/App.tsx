import { useCallback } from "react";

import { ToolCard } from "./components/ToolCard";

type MicrosoftTeamsClient = {
  app: {
    initialize: () => Promise<void>;
    openLink: (url: string) => Promise<void>;
  };
  dialog: {
    url: {
      open: (options: {
        url: string;
        size: { height: number; width: number };
        title: string;
      }) => Promise<void>;
    };
  };
};

type Tool = {
  key: string;
  title: string;
  description: string;
  badge: string;
  accentColor: string;
  icon: string;
  url: string;
  dialogUrl?: string;
  dialogSize?: { width: number; height: number };
  guideUrl?: string;
  launchPreference?: "dialog" | "link" | "browser";
};

const DEFAULT_DIALOG_SIZE = { width: 1600, height: 1100 };

const tools: Tool[] = [
  {
    key: "Nasma",
    title: "Nasma: The P&C Bot",
    description:
      "Intelligent AI assistant designed to streamline P&C processes and employee interactions. Provides instant responses to common workplace queries and automates routine P&C tasks.",
    badge: "For everyone",
    accentColor: "#FF6666",
    icon: "bot",
    url: "https://nasma-production.up.railway.app/",
    dialogUrl: "https://nasma-production.up.railway.app/",
    dialogSize: { width: 1600, height: 1200 }
  },
  {
    key: "PrezlabBrain",
    title: "The Prezlab Brain",
    description:
      "AI-powered creative assistant that analyzes client presentations and provides strategic recommendations aligned with Prezlab's creative standards and best practices.",
    badge: "For creatives",
    accentColor: "#805AF9",
    icon: "brain",
    url: "https://prezlab-brain.netlify.app/"
  },
  {
    key: "TMS",
    title: "Traffic Management System (TMS)",
    description:
      "Comprehensive project routing and assignment platform that optimizes workflow distribution and ensures efficient project completion.",
    badge: "For client success",
    accentColor: "#4EF4A8",
    icon: "route",
    url: "https://prez-tms.up.railway.app/"
  },
  {
    key: "UtilizationDashboard",
    title: "Utilization Dashboard",
    description:
      "Comprehensive insights into team capacity, workload distribution, and resource optimization to maximize productivity.",
    badge: "For managers",
    accentColor: "#FFC952",
    icon: "chart",
    url: "https://utdashboard-production.up.railway.app/"
  },
  {
    key: "LeadHub",
    title: "Lead Hub",
    description:
      "Intelligent lead management system that automates enrichment, qualification, and CRM integration for business development.",
    badge: "For BD & marketing",
    accentColor: "#0EDCFB",
    icon: "target",
    url: "https://lead-automation-system.onrender.com/"
  }
];

export default function App() {
  const launchTool = useCallback(async (tool: Tool) => {
    const targetUrl = tool.dialogUrl ?? tool.url;

    if (tool.launchPreference === "browser") {
      window.open(tool.url, "_blank", "noopener,noreferrer");
      return;
    }

    const teams = (window as typeof window & { microsoftTeams?: MicrosoftTeamsClient }).microsoftTeams;

    if (!teams) {
      window.open(tool.url, "_blank", "noopener,noreferrer");
      return;
    }

    try {
      // Initialize Teams SDK first
      await teams.app.initialize();

      if (tool.launchPreference === "link") {
        await teams.app.openLink(targetUrl);
        return;
      }

      // Default: open in dialog (iframe) - maximize size
      await teams.dialog.url.open({
        url: targetUrl,
        size: tool.dialogSize ?? DEFAULT_DIALOG_SIZE,
        title: tool.title
      });
      return;
    } catch (error) {
      console.error("Failed to launch in Teams:", error);
      try {
        await teams.app.openLink(tool.url);
        return;
      } catch (linkError) {
        console.error("Teams openLink fallback failed:", linkError);
        window.open(tool.url, "_blank", "noopener,noreferrer");
      }
    }
  }, []);

  const launchInBrowser = useCallback((tool: Tool) => {
    window.open(tool.url, "_blank", "noopener,noreferrer");
  }, []);

  const learnHowToUse = useCallback((tool: Tool) => {
    const destination = tool.guideUrl ?? tool.url;
    window.open(destination, "_blank", "noopener,noreferrer");
  }, []);

  return (
    <div className="layout">
      <header className="hero" aria-labelledby="hero-title">
        <div className="hero-top">
          <div className="hero-wordmark" aria-label="PREZ.AI">
            PREZ<span className="accent-dot">.</span>AI
          </div>
          <a
            className="hero-feedback"
            href="https://forms.office.com/Pages/ResponsePage.aspx?id=VUu9FkZ0GkaCAp3lh3RgisMasIxrLtFDu7EJ7Tl-U8FUNFlLQ0dXN0szSkRORDlNRVhMVTZST1VJTy4u"
            target="_blank"
            rel="noreferrer"
          >
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
            </svg>
            Feedback
          </a>
        </div>
        <h1 id="hero-title">Your gateway to intelligent workflows</h1>
        <p className="hero-subtitle">
          Discover AI copilots tailored to each Prezlab team&mdash;explore the toolkit built for your day-to-day.
        </p>
      </header>

      <main className="tools-grid" aria-label="PREZ.AI tools">
        {tools.map((tool) => (
          <ToolCard
            key={tool.key}
            toolKey={tool.key}
            title={tool.title}
            description={tool.description}
            badge={tool.badge}
            accentColor={tool.accentColor}
            icon={tool.icon}
            onLaunch={() => launchTool(tool)}
            onLaunchInBrowser={() => launchInBrowser(tool)}
            onLearnHowToUse={() => learnHowToUse(tool)}
          />
        ))}
      </main>

      <footer className="footer">
        <p>
          Powered by <span className="footer-accent">PREZ<span className="accent-dot">.</span>AI</span> &mdash; AC 2025
        </p>
      </footer>
    </div>
  );
}
