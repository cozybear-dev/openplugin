import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle
} from "@fluentui/react-components";

export function RestoreDialog(props: {
  open: boolean;
  onCancel: () => void;
  onConfirm: () => void;
}) {
  return (
    <Dialog
      modalType="alert"
      open={props.open}
      onOpenChange={(_, data) => {
        if (!data.open) props.onCancel();
      }}
    >
      <DialogSurface className="op-restore-dialog">
        <DialogBody>
          <DialogTitle>Restore to this message?</DialogTitle>
          <DialogContent>
            Chat after this message will be removed. OpenPlugin edits made after it are undone when
            possible.
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={props.onCancel}>
              Cancel
            </Button>
            <Button appearance="primary" onClick={props.onConfirm}>
              Restore
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
