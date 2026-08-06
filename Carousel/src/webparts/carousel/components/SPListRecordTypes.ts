export type carouselSlideRecord = {
  Id: string;
  SlideOrder?: number;
  Title?: string;
  ClickLink?: string;
  LinkTarget?: string;
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
