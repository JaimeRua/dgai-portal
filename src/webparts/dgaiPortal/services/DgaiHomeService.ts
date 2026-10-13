import type { SPFI } from '@pnp/sp';

import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import type { IDgaiHomeItem } from '../models/IDgaiHomeItem';

interface IHomeListItem {
  Title?: string;
  Codigo?: string;
  Seccion?: string;
  OrdenSeccion?: number;
  Orden?: number;

  Url?: string | {
    Url?: string;
  };

  TipoE?: 'Biblioteca' | 'Lista' | 'Página' | 'Agenda';
  Icono?: string;
  Activo?: boolean;
}

export class DgaiHomeService {

  private readonly _sp: SPFI;
  private readonly _configurationListName: string;

  public constructor(
    sp: SPFI,
    configurationListName = 'DGAI_ConfiguracionInicio'
  ) {
    this._sp = sp;
    this._configurationListName = configurationListName;
  }

  private _getUrl(
    value: string | { Url?: string } | undefined
  ): string {

    if (!value) {
      return '#';
    }

    if (typeof value === 'string') {
      return value;
    }

    return value.Url || '#';
  }

  public async getHomeItems(): Promise<IDgaiHomeItem[]> {

    if (!this._configurationListName?.trim()) {
      throw new Error(
        'No se ha configurado la lista de inicio del portal.'
      );
    }

    const items: IHomeListItem[] =
      await this._sp.web.lists
        .getByTitle(this._configurationListName)
        .items
        .select(
          'Title',
          'Codigo',
          'Seccion',
          'OrdenSeccion',
          'Orden',
          'Url',
          'TipoE',
          'Icono',
          'Activo'
        )
        .filter('Activo eq 1')
        .orderBy('OrdenSeccion', true)
        .orderBy('Orden', true)();

    return items.map((item) => ({
      id: item.Codigo || '',
      title: item.Title || '',
      section: item.Seccion || '',
      sectionOrder: item.OrdenSeccion || 0,
      order: item.Orden || 0,
      url: this._getUrl(item.Url),
      type: item.TipoE || 'Página',
      icon: item.Icono || '',
      active: item.Activo ?? true
    }));
  }
}