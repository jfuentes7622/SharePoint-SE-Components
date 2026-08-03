import { PersonObjectModel } from './PersonObject';

export type PersonGroupModel = {
    groupTitle: string;
    personList: Array<PersonObjectModel>;
};
