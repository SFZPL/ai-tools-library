
// Tool URLs mapping
const toolUrls = {
  'Nasma': 'https://en.wikipedia.org/wiki/Artificial_intelligence', // Test URL
  'PrezBot': 'https://example.com/prezbot',
  'TMS': 'https://example.com/tms',
  'CashCollection': 'https://example.com/cashcollection',
  'UtilizationDashboard': 'https://example.com/utilization',
  'LeadEnrichment': 'https://example.com/leadenrichment'
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
