export interface ITravelSiteItem {
  Id?: number;
  Title: string; 
  Place: string;
  FamousFood: string;
  Price: 'Budget-friendly' | 'Premium Experience' | 'No cost' | 'Mid-range';
}