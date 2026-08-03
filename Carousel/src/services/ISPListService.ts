import {carouselSlideRecord } from '../webparts/carousel/components/SPListRecordTypes';

export interface ISPListService {
 
    GetSlideData(slideLib:string): Promise<Array<carouselSlideRecord>>;
    
}
