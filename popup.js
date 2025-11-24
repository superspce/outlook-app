// Popup script for Outlook Auto Attach extension

// Check if there's a pending file on load
chrome.runtime.sendMessage({ action: 'getPendingFile' }, (response) => {
  if (response && response.filePath) {
    showConfirmation(response.filePath, response.downloadId);
  } else {
    // Show waiting message
    document.getElementById('waiting').classList.remove('hidden');
    document.getElementById('confirmation').classList.add('hidden');
  }
});

function showConfirmation(filePath, downloadId) {
  document.getElementById('waiting').classList.add('hidden');
  document.getElementById('confirmation').classList.remove('hidden');
  
  // Display filename
  const filenameElement = document.getElementById('filename');
  if (filePath) {
    const filename = filePath.split('/').pop().split('\\').pop();
    filenameElement.textContent = filename;
  } else {
    filenameElement.textContent = 'Unknown file';
  }

  // Handle cancel button
  document.getElementById('cancelBtn').addEventListener('click', () => {
    chrome.runtime.sendMessage({
      action: 'cancelOutlook',
      downloadId: downloadId
    });
    window.close();
  });

  // Handle confirm button
  document.getElementById('confirmBtn').addEventListener('click', () => {
    chrome.runtime.sendMessage({
      action: 'confirmOutlook',
      filePath: filePath,
      downloadId: downloadId
    }, (response) => {
      if (chrome.runtime.lastError) {
        console.error('Error:', chrome.runtime.lastError);
      }
      window.close();
    });
  });
}

// Listen for messages from background script
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'showConfirmation') {
    showConfirmation(request.filePath, request.downloadId);
    sendResponse({ success: true });
  }
});

