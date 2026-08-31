type CanvasContextMenuProps = {
  x: number
  y: number
  canUndo: boolean
  canCut: boolean
  canPaste: boolean
  showProperties?: boolean
  onUndo: () => void
  onCut: () => void
  onPaste: () => void
  onProperties?: () => void
  onClose: () => void
}

export const CanvasContextMenu = ({
  x,
  y,
  canUndo,
  canCut,
  canPaste,
  showProperties,
  onUndo,
  onCut,
  onPaste,
  onProperties,
  onClose,
}: CanvasContextMenuProps) => {
  const left = Math.min(x, Math.max(8, window.innerWidth - 220))
  const top = Math.min(y, Math.max(8, window.innerHeight - 220))

  return (
    <div
      className="canvas-menu"
      style={{ left, top }}
      role="menu"
      aria-label="Canvas menu"
      onMouseDown={(event) => event.stopPropagation()}
      onContextMenu={(event) => event.preventDefault()}
    >
      <button
        type="button"
        role="menuitem"
        disabled={!canUndo}
        onClick={() => {
          onUndo()
          onClose()
        }}
      >
        <span>Undo</span>
        <kbd>Ctrl+Z</kbd>
      </button>
      <button
        type="button"
        role="menuitem"
        disabled={!canCut}
        onClick={() => {
          onCut()
          onClose()
        }}
      >
        <span>Cut</span>
        <kbd>Ctrl+X</kbd>
      </button>
      <button
        type="button"
        role="menuitem"
        disabled={!canPaste}
        onClick={() => {
          onPaste()
          onClose()
        }}
      >
        <span>Paste</span>
        <kbd>Ctrl+V</kbd>
      </button>
      {showProperties && onProperties && (
        <>
          <div className="canvas-menu__sep" />
          <button
            type="button"
            role="menuitem"
            onClick={() => {
              onProperties()
              onClose()
            }}
          >
            <span>Properties</span>
          </button>
        </>
      )}
    </div>
  )
}
