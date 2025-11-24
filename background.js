// Service worker startup
console.log('Outlook Auto Attach extension service worker started');
console.log('Extension ID:', chrome.runtime.id);
console.log('Manifest version:', chrome.runtime.getManifest().version);

// Keep service worker alive and verify it's working
chrome.runtime.onInstalled.addListener(() => {
  console.log('Extension installed/updated');
});

// Add alarm to keep service worker alive (optional, for testing)
chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === 'keep-alive') {
    console.log('Service worker keep-alive ping');
  }
});

// Create a keep-alive alarm (every 30 seconds)
chrome.alarms.create('keep-alive', { periodInMinutes: 0.5 });

// Function to check if filename matches criteria
function shouldProcessFile(filePath) {
  if (!filePath) return false;
  
  // Extract filename from path
  const filename = filePath.split('/').pop().split('\\').pop();
  
  // Normalize to lowercase for comparison
  const filenameLower = filename.toLowerCase();
  
  // Check if filename includes "Orderbekräftelse" (case-insensitive) or "1000322"
  const includesOrderbekraeftelse = filenameLower.includes('orderbekräftelse') || 
                                     filenameLower.includes('orderbekr');
  const includesOrderNumber = filenameLower.includes('1000322');
  
  const shouldProcess = includesOrderbekraeftelse || includesOrderNumber;
  
  console.log('📋 File filter check:', {
    originalFilename: filename,
    filenameLower: filenameLower,
    includesOrderbekraeftelse: includesOrderbekraeftelse,
    includesOrderNumber: includesOrderNumber,
    shouldProcess: shouldProcess
  });
  
  return shouldProcess;
}

// Listen for download created events
chrome.downloads.onCreated.addListener((downloadItem) => {
  console.log('========== DOWNLOAD CREATED ==========');
  console.log('Download ID:', downloadItem.id);
  console.log('Filename:', downloadItem.filename);
  console.log('URL:', downloadItem.url);
  console.log('State:', downloadItem.state);
  console.log('======================================');
  
  // Handle downloads that are already complete when created
  // (some downloads complete so fast they're already done)
  if (downloadItem.state === 'complete' && downloadItem.error === undefined) {
    const filePath = downloadItem.filename;
    if (filePath && shouldProcessFile(filePath)) {
          console.log('🎉 Download already complete when created!');
          console.log('📁 File path:', filePath);
          console.log('✅ File matches filter - showing confirmation...');
          showConfirmationDialog(filePath, downloadItem.id);
    } else if (filePath) {
          console.log('⏭️ File does not match filter criteria - skipping:', filePath);
    }
  }
});

// Listen for download completion events
chrome.downloads.onChanged.addListener((downloadDelta) => {
  console.log('========== DOWNLOAD CHANGED ==========');
  console.log('Download ID:', downloadDelta.id);
  console.log('State change:', downloadDelta.state);
  console.log('Full delta:', JSON.stringify(downloadDelta, null, 2));
  console.log('======================================');
  
  // Check if the download has completed successfully
  if (downloadDelta.state && downloadDelta.state.current === 'complete') {
    console.log('✅ Download state changed to COMPLETE for ID:', downloadDelta.id);
    
    // Get the download item to retrieve the file path
    chrome.downloads.search({ id: downloadDelta.id }, (downloads) => {
      console.log('Search results for download ID:', downloadDelta.id, '- Found:', downloads.length, 'items');
      
      if (downloads.length > 0) {
        const download = downloads[0];
        console.log('========== DOWNLOAD DETAILS ==========');
        console.log('ID:', download.id);
        console.log('Filename:', download.filename);
        console.log('State:', download.state);
        console.log('Error:', download.error);
        console.log('Total bytes:', download.totalBytes);
        console.log('File size:', download.fileSize);
        console.log('======================================');
        
        // Verify download completed successfully (not interrupted or failed)
        if (download.state === 'complete' && download.error === undefined) {
          const filePath = download.filename;
          
          console.log('🎉 Download completed successfully!');
          console.log('📁 File path:', filePath);
          
          // Check if file matches filter criteria before processing
          if (shouldProcessFile(filePath)) {
            console.log('✅ File matches filter - showing confirmation...');
            // Show confirmation dialog before opening Outlook
            showConfirmationDialog(filePath, download.id);
          } else {
            console.log('⏭️ File does not match filter criteria - skipping:', filePath);
          }
        } else {
          console.error('❌ Download did not complete successfully:', {
            state: download.state,
            error: download.error,
            filename: download.filename
          });
        }
      } else {
        console.error('❌ Download not found for ID:', downloadDelta.id);
      }
    });
  } else if (downloadDelta.state) {
    console.log('⏳ Download state:', downloadDelta.state.current, 'for ID:', downloadDelta.id);
  }
});

// Store pending files waiting for user confirmation
const pendingFiles = new Map();

// Function to show confirmation dialog
function showConfirmationDialog(filePath, downloadId) {
  console.log('Showing confirmation dialog for:', filePath);
  
  // Store the file path for when user confirms
  pendingFiles.set(downloadId, filePath);
  
  // Set badge on extension icon to notify user
  chrome.action.setBadgeText({ text: '1' });
  chrome.action.setBadgeBackgroundColor({ color: '#0078d4' });
  chrome.action.setTitle({ title: 'Click to confirm sending file via Outlook' });
  
  // Try to open the popup programmatically (may not work in all cases)
  // User will see the badge and can click the extension icon
  chrome.action.openPopup(() => {
    if (chrome.runtime.lastError) {
      console.log('Cannot auto-open popup (user must click extension icon):', chrome.runtime.lastError.message);
      // Send message to popup if it's already open
      chrome.runtime.sendMessage({
        action: 'showConfirmation',
        filePath: filePath,
        downloadId: downloadId
      }).catch(() => {
        // Popup not open, user will see badge and click icon
      });
    }
  });
  
  // Also show a notification to guide user
  chrome.notifications.create({
    type: 'basic',
    iconUrl: 'icons/icon48.png',
    title: 'Outlook Auto Attach',
    message: 'Click the extension icon to confirm sending file via Outlook',
    requireInteraction: false
  }, (notificationId) => {
    // Auto-clear notification after 5 seconds
    setTimeout(() => {
      chrome.notifications.clear(notificationId);
    }, 5000);
  });
}

// Listen for messages from popup
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'getPendingFile') {
    // Return the first pending file
    const entries = Array.from(pendingFiles.entries());
    if (entries.length > 0) {
      const [downloadId, filePath] = entries[0];
      sendResponse({ filePath: filePath, downloadId: downloadId });
    } else {
      sendResponse({ filePath: null });
    }
    return true;
  }
  
  if (request.action === 'confirmOutlook') {
    const filePath = request.filePath || pendingFiles.get(request.downloadId);
    
    if (filePath) {
      console.log('User confirmed - opening Outlook with:', filePath);
      // Clear badge
      chrome.action.setBadgeText({ text: '' });
      chrome.action.setTitle({ title: 'Outlook Auto Attach' });
      // Remove from pending
      if (request.downloadId) {
        pendingFiles.delete(request.downloadId);
      }
      // Send to server to open Outlook
      sendToServer(filePath);
      sendResponse({ success: true });
    } else {
      console.error('File path not found for download ID:', request.downloadId);
      sendResponse({ success: false, error: 'File path not found' });
    }
    return true;
  }
  
  if (request.action === 'cancelOutlook') {
    console.log('User cancelled');
    // Clear badge
    chrome.action.setBadgeText({ text: '' });
    chrome.action.setTitle({ title: 'Outlook Auto Attach' });
    // Remove from pending
    if (request.downloadId) {
      pendingFiles.delete(request.downloadId);
    }
    sendResponse({ success: true });
    return true;
  }
});

// Function to send file path to local server
function sendToServer(filePath) {
  console.log('Attempting to send file to local server:', filePath);
  
  const serverUrl = 'http://localhost:8765/attach';
  
  // Send POST request to local server
  fetch(serverUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ 
      filePath: filePath
    })
  })
  .then(response => response.json())
  .then(data => {
    console.log('Server response:', data);
    
    if (data.success) {
      // Show success notification
      chrome.notifications.create({
        type: 'basic',
        iconUrl: 'icons/icon48.png',
        title: 'Outlook Auto Attach',
        message: 'Opening Outlook with attached file...'
      }, (notificationId) => {
        console.log('Success notification shown:', notificationId);
      });
    } else {
      // Show error notification
      chrome.notifications.create({
        type: 'basic',
        iconUrl: 'icons/icon48.png',
        title: 'Outlook Auto Attach',
        message: `Failed to open Outlook: ${data.message || 'Unknown error'}`
      }, (notificationId) => {
        console.log('Error notification shown:', notificationId);
      });
    }
  })
  .catch(error => {
    console.error('Error communicating with server:', error);
    
    // Show error notification
    chrome.notifications.create({
      type: 'basic',
      iconUrl: 'icons/icon48.png',
      title: 'Outlook Auto Attach',
      message: 'Failed to connect to local server. Make sure the server is running.'
    }, (notificationId) => {
      console.log('Error notification shown:', notificationId);
    });
  });
}

