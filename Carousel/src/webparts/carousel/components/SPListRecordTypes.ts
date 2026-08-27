export type carouselSlideRecord = {
  [key: string]: any;
  Id: string;
  SlideOrder?: number;
  Title?: string;
  Description?: string;
  ClickLink?: string | { Url?: string; Description?: string; };
  LinkTarget?: string | { Url?: string; Description?: string; };
  FileRef: string;
  FileLeafRef?: string;
  FSObjType?: number;
  Display?: string | boolean;
  Expiration?: string;
  StartDate?: string;
  File?: {
    ServerRelativeUrl?: string;
    Name?: string;
  };
};
