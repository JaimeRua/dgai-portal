export interface IDgaiPortalItem {
  id: string;
  order: number;
  title: string;
  description: string;
  url: string;
  linkText: string;
  type: 'Área' | 'Acceso Rápido';
  active: boolean;
}