import SlideItemModel from './SlideItem';
import RecordSvc from '../../../services/RecordService';
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICarouselProps {
  carouselBackgroundColor: string;
  carouselWidth: number;
  carouselHeight:number;
  carouselSlideInterval: string;
  carouselSlideItems: Array<SlideItemModel>;
  carouselTransitionInterval: string;
  carouselSlideLibrary: string;
  spfxContext: WebPartContext;
  recSvc: RecordSvc;
}
