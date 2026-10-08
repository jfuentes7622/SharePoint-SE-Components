export interface IScriptEditorWebPartProps {
  script: string;
  title: string;
  removePadding: boolean;
  spPageContextInfo: boolean;
  excludeFromTabs: boolean;
  overrideCssUrl?: string;
  propPaneHandle: { open: () => void };
}
