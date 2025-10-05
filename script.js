
// Tool URLs mapping
const toolUrls = {
  'Nasma': 'https://en.wikipedia.org/wiki/Artificial_intelligence', // Test URL
  'PrezlabBrain': 'https://example.com/prezlab-brain',
  'TMS': 'https://example.com/tms',
  'UtilizationDashboard': 'https://example.com/utilization',
  'LeadHub': 'https://example.com/lead-hub'
};

function openTool(toolName) {
  const url = toolUrls[toolName];

  // Check if running in Teams
  if (typeof microsoftTeams !== 'undefined') {
    microsoftTeams.app.initialize().then(() => {
      // Open in a stage view (modal) inside Teams
      microsoftTeams.app.openLink({
        url: url,
        targetElementBounds: null
      }).catch(() => {
        // If stage view fails, try dialog
        microsoftTeams.dialog.url.open({
          url: url,
          size: {
            height: 600,
            width: 800
          },
          title: toolName
        }).catch(() => {
          // Final fallback: open in browser
          window.open(url, '_blank');
        });
      });
    });
  } else {
    // Fallback for testing outside Teams
    window.open(url, '_blank');
  }
}
