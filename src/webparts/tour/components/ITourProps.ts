export interface ITourProps {
  description: string;
  actionValue: string;
  icon: string;
  collectionData: any[];
  onClose: () => void;
  editMode: boolean;
  isFullWidth: boolean;
}
