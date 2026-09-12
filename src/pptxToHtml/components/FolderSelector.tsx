import React from 'react'
import './FolderSelector.css'

interface FolderSelectorProps {
  onFolderSelect: (folder: string) => void
  selectedFolder: string
}

const FolderSelector: React.FC<FolderSelectorProps> = ({ onFolderSelect }) => {
  const handleAcknowledge = () => {
    // User acknowledges they understand the download location
    onFolderSelect('Downloads')
  }

  return (
    <div className="folder-selector">
      <div className="folder-selector-label">
        <h4>Download Location:</h4>
        <p>The converted HTML file will be saved to your Downloads folder</p>
      </div>
      
      <div className="download-info">
        <div className="info-box">
          <span className="info-icon">ℹ️</span>
          <div className="info-content">
            <p><strong>Browser Security Notice:</strong></p>
            <p>Due to browser security restrictions, files cannot be saved to custom locations. 
               The converted HTML file will be automatically downloaded to your default Downloads folder.</p>
          </div>
        </div>
        
        <button 
          type="button" 
          onClick={handleAcknowledge}
          className="acknowledge-button"
        >
          ✓ I Understand
        </button>
      </div>
    </div>
  )
}

export default FolderSelector
