export interface IDgaiHomeItem {
  id: string;
  title: string;
  section: string;
  sectionOrder: number;
  order: number;
  url: string;
  type: 'Biblioteca' | 'Lista' | 'Página' | 'Agenda';
  icon?: string;
  active: boolean;
}