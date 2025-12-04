// This file handles commands from the ribbon buttons
// Initialize Office.js when the commands page loads

// Office.js is loaded via script tag in commands.html
// Access Office via window to avoid TypeScript type checking issues
declare const Office: {
  onReady: (callback: (info: { host: string; platform: string }) => void) => void;
  HostType: {
    PowerPoint: string;
    Excel: string;
    Word: string;
    Outlook: string;
  };
};

// Use type assertion to access Office from window
const OfficeAPI = (window as any).Office as typeof Office;

if (typeof OfficeAPI !== 'undefined') {
  OfficeAPI.onReady((info) => {
    if (info.host === OfficeAPI.HostType.PowerPoint) {
      console.log('Commands loaded successfully for PowerPoint');
    }
  });
}

// If you need to handle specific ribbon button actions in the future,
// you can add them here. For now, the ShowTaskpane action is handled automatically
// by the manifest.xml configuration.
