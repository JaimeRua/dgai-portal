import * as React from 'react';

import styles from './DgaiPortal.module.scss';

import type { IDgaiPortalProps } from './IDgaiPortalProps';
import type { IDgaiHomeItem } from '../models/IDgaiHomeItem';

import dgaiLogo from '../assets/dgai-logo.png';
import { getSafePortalUrl } from '../../../shared/security/portalUrl';

interface IDgaiPortalState {
  homeItems: IDgaiHomeItem[];
  loading: boolean;
  error?: string;
}

export default class DgaiPortal extends React.Component<
  IDgaiPortalProps,
  IDgaiPortalState
> {

  public constructor(props: IDgaiPortalProps) {
    super(props);

    this.state = {
      homeItems: [],
      loading: true
    };
  }

  public async componentDidMount(): Promise<void> {

    try {

      const homeItems =
        await this.props.homeService.getHomeItems();

      console.log(
        'DGAI - home items:',
        homeItems
      );

      this.setState({
        homeItems,
        loading: false
      });

    } catch (error) {

      console.error(
        '[DgaiPortal] Error cargando datos:',
        error
      );

      this.setState({
        loading: false,
        error:
          error instanceof Error
            ? error.message
            : 'Error al cargar la configuración del portal.'
      });
    }
  }

  public render(): React.ReactElement<IDgaiPortalProps> {

    if (this.state.loading) {
      return (
        <section className={styles.dgaiPortal}>
          <p>
            Cargando configuración del portal...
          </p>
        </section>
      );
    }

    if (this.state.error) {
      return (
        <section className={styles.dgaiPortal}>
          <p>
            Error al cargar la configuración:{' '}
            {this.state.error}
          </p>
        </section>
      );
    }

    const homeItems =
      [...this.state.homeItems]
        .sort((a, b) => {

          if (a.sectionOrder !== b.sectionOrder) {
            return a.sectionOrder - b.sectionOrder;
          }

          return a.order - b.order;
        });

    const sections =
      homeItems.reduce<
        Record<string, IDgaiHomeItem[]>
      >((groups, item) => {

        if (!groups[item.section]) {
          groups[item.section] = [];
        }

        groups[item.section].push(item);

        return groups;

      }, {});

    return (

      <section className={styles.dgaiPortal}>

        <div className={styles.content}>

          <div className={styles.homeLayout}>

  <div className={styles.homeSections}>

    {Object.keys(sections).map(
      (sectionName: string) => {

        const sectionItems: IDgaiHomeItem[] =
          sections[sectionName];

        return (
          <section
            key={sectionName}
            className={styles.homeSection}
          >

            <h2
              className={styles.homeSectionTitle}
            >
              {sectionName}
            </h2>

            <div className={styles.homeLinks}>

              {sectionItems.map(
                (item: IDgaiHomeItem) => (

                  <a
                    key={item.id}
                    href={getSafePortalUrl(item.url)}
                    target="_blank"
                    rel="noopener noreferrer"
                    className={styles.homeLink}
                  >

                    <span
                      className={styles.homeLinkIcon}
                      aria-hidden="true"
                    >
                      {this._getItemIcon(item)}
                    </span>

                    <span>
                      {item.title}
                    </span>

                  </a>

                )
              )}

            </div>

          </section>
        );
      }
    )}

  </div>

  <aside className={styles.homeBrand}>
  <div className={styles.logoContainer}>
    <img
      src={dgaiLogo}
      alt="Dirección de Gestión y Análisis Institucional"
      className={styles.dgaiLogo}
    />
  </div>
</aside>

</div>

        </div>

      </section>
    );
  }

  private _getItemIcon(
    item: IDgaiHomeItem
  ): string {

    switch (item.type) {

      case 'Agenda':
        return '📅';

      case 'Biblioteca':
        return '📁';

      case 'Lista':
        return '📋';

      case 'Página':
        return '🔗';

      default:
        return '🔗';
    }
  }
}