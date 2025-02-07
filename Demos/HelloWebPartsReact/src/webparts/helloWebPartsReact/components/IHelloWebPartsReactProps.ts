import {
  ButtonClickedCallback,
  IListItem
} from '../../../models';

export interface IHelloWebPartsReactProps {
  spListItems: IListItem[];
  onGetListItems: ButtonClickedCallback;
  description: string;
  listName: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
