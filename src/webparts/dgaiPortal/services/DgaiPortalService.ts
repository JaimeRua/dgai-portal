import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import type { IDgaiPortalItem } from '../models/IDgaiPortalItem';

interface IPortalListItem {
  Title?: string;
  Codigo?: string;
  Descripcion?: string;
  Url?: {
    Url?: string;
  };
  TextoEnlace?: string;
  Orden?: number;
  TipoEnlace?: 'Área' | 'Acceso Rápido';
  Activo?: boolean;
}

export class DgaiPortalService {

  private readonly _sp: SPFI;
  private readonly _configurationListName: string;

  constructor(
    sp: SPFI,
    configurationListName: string
  ) {
    this._sp = sp;
    this._configurationListName = configurationListName;
  }

  public async getPortalItems(): Promise<IDgaiPortalItem[]> {

    if (!this._configurationListName?.trim()) {
      throw new Error(
        'No se ha configurado una lista para el portal.'
      );
    }

    const items: IPortalListItem[] =
      await this._sp.web.lists
        .getByTitle(this._configurationListName)
        .items
        .select(
          'Title',
          'Codigo',
          'Descripcion',
          'Url',
          'TextoEnlace',
          'Orden',
          'TipoEnlace',
          'Activo'
        )
        .filter('Activo eq 1')
        .orderBy('Orden', true)();

    return items.map((item) => ({
      id: item.Codigo || '',
      order: item.Orden || 0,
      title: item.Title || '',
      description: item.Descripcion || '',
      url: item.Url?.Url || '#',
      linkText: item.TextoEnlace || item.Title || '',
      type: item.TipoEnlace || 'Área',
      active: item.Activo ?? true
    }));
  }
}