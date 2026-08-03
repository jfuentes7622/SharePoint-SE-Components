import SlideItemModel from '../webparts/carousel/components/SlideItem';

export interface IRecordService {   

    GetSlides(slideLib: string): Promise<Array<SlideItemModel>>;
    
}
