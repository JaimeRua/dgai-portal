import type { SPFI } from '@pnp/sp';

import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import type { IDgaiLayoutItem } from '../models/IDgaiLayoutItem';

interface ILayoutListItem {
  Title?: string;
  Codigo?: string;

  Url?: string | {
    Url?: string;
  };

  Orden?: number;
  Grupo?: string;
  Activo?: boolean;
  SeparadorDespues?: boolean;
}

export class DgaiLayoutService {

  private readonly _sp: SPFI;

  public constructor(sp: SPFI) {
    this._sp = sp;
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

  public async getNavigationItems():
    Promise<IDgaiLayoutItem[]> {

    const items: ILayoutListItem[] =
      await this._sp.web.lists
        .getByTitle('DGAI_ConfiguracionLayout')
        .items
        .select(
          'Title',
          'Codigo',
          'Url',
          'Orden',
          'Grupo',
          'Activo',
          'SeparadorDespues'
        )
        .filter('Activo eq 1')
        .orderBy('Orden', true)();

    return items.map((item) => ({
      id: item.Codigo || '',
      title: item.Title || '',
      url: this._getUrl(item.Url),
      order: item.Orden || 0,
      group: item.Grupo || '',
      active: item.Activo ?? true,
      separatorAfter: item.SeparadorDespues ?? false
    }));
  }
}