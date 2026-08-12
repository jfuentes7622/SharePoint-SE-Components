/**
 * Form Templates
 */

import { FormSchema } from '../formEngine/core/types';

export const BlankTemplate: FormSchema = {
  id: 'blank-form',
  name: 'Blank Form',
  mode: 'new',
  listName: '',
  steps: [{ id: 'step1', title: 'Container 1', fields: [] }],
};
