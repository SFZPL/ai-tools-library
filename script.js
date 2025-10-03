
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
      // Try to open in Teams (will open in iframe if site allows it)
      microsoftTeams.app.openLink(url);
    });
  } else {
    // Fallback for testing outside Teams
    window.open(url, '_blank');
  }
}
