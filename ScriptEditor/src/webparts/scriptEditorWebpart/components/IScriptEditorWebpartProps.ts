export interface IScriptEditorWebPartProps {
  script: string;
  title: string;
  removePadding: boolean;
  spPageContextInfo: boolean;
  propPaneHandle: { open: () => void };
}
