const toolCatalog = {
  Nasma: {
    name: "Nasma - The P&C Bot",
    url: "https://en.wikipedia.org/wiki/Artificial_intelligence"
  },
  PrezlabBrain: {
    name: "The Prezlab Brain",
    url: "https://example.com/prezlab-brain"
  },
  TMS: {
    name: "Traffic Management System (TMS)",
    url: "https://example.com/tms"
  },
  UtilizationDashboard: {
    name: "Utilization Dashboard",
    url: "https://example.com/utilization"
  },
  LeadHub: {
    name: "Lead Hub",
    url: "https://example.com/lead-hub"
  }
};

async function launchTool(toolKey) {
  const tool = toolCatalog[toolKey];
  if (!tool) {
    console.warn(`Unknown tool key: ${toolKey}`);
    return;
  }

  if (typeof microsoftTeams === "undefined") {
    window.open(tool.url, "_blank", "noopener,noreferrer");
    return;
  }

  try {
    await microsoftTeams.app.initialize();
    await microsoftTeams.app.openLink(tool.url);
  } catch (stageViewError) {
    try {
      await microsoftTeams.dialog.url.open({
        url: tool.url,
        size: { height: 1000, width: 1400 },
        title: tool.name
      });
    } catch (dialogError) {
      window.open(tool.url, "_blank", "noopener,noreferrer");
    }
  }
}

function launchInBrowser(toolKey) {
  const tool = toolCatalog[toolKey];
  if (!tool) {
    console.warn(`Unknown tool key: ${toolKey}`);
    return;
  }
  window.open(tool.url, "_blank", "noopener,noreferrer");
}

document.addEventListener("DOMContentLoaded", () => {
  document.querySelectorAll('[data-action="launch"]').forEach((button) => {
    button.addEventListener("click", (event) => {
      const toolKey = event.currentTarget.getAttribute("data-tool-key");
      launchTool(toolKey);
    });
  });

  document.querySelectorAll('[data-action="browser"]').forEach((button) => {
    button.addEventListener("click", (event) => {
      const toolKey = event.currentTarget.getAttribute("data-tool-key");
      launchInBrowser(toolKey);
      event.stopPropagation();
    });
  });
});
