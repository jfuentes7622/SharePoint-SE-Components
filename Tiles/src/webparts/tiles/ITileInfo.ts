export interface ITileInfo {
  title: string;
  description: string;
  url: string;
  icon: string;
  target: LinkTarget;
  /** Per-tile background colour override (overrides global backgroundColor). */
  color?: string;
  imageUrl?: string;
  /** When true the image fills the entire tile and no text is shown. */
  imageOnly?: boolean;
  imagePosition?: 'top' | 'bottom' | 'left' | 'right';
  /** Optional image shown only on rollover/hover overlay. */
  hoverImageUrl?: string;
  /** When true on hover, the image fills the overlay and no text is shown. */
  hoverImageOnly?: boolean;
  /** Position of hover image when hoverImageOnly is false. */
  hoverImagePosition?: 'top' | 'bottom' | 'left' | 'right';
}

export enum LinkTarget {
  parent = '',
  blank = '_blank'
}
