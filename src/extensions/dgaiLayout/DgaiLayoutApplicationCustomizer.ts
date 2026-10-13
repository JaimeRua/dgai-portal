import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import styles from './DgaiLayout.module.scss';

import { spfi, SPFx } from '@pnp/sp';
import type { SPFI } from '@pnp/sp';

import { DgaiLayoutService } from './services/DgaiLayoutService';
import type { IDgaiLayoutItem } from './models/IDgaiLayoutItem';

export interface IDgaiLayoutApplicationCustomizerProperties {
  testMessage?: string;
}

export default class DgaiLayoutApplicationCustomizer
  extends BaseApplicationCustomizer<IDgaiLayoutApplicationCustomizerProperties> {

  private _topPlaceholder?: PlaceholderContent;
  private _contentElement?: HTMLElement;
  private _originalMarginLeft?: string;
  private _originalWidth?: string;
  private _siteHeaderElement?: HTMLElement;
  private _originalSiteHeaderMarginLeft?: string;
  private _originalSiteHeaderWidth?: string;
  private _originalSiteHeaderDisplay?: string;
  private _appBarElement?: HTMLElement;
  private _originalAppBarDisplay?: string;
  private _appBarObserver?: MutationObserver;
  private _siteHeaderObserver?: MutationObserver;
  private _suiteNavElement?: HTMLElement;
  private _originalSuiteNavDisplay?: string;
  private _suiteNavObserver?: MutationObserver;
  private _sp!: SPFI;
  private _layoutService!: DgaiLayoutService;
  private _navigationItems: IDgaiLayoutItem[] = [];
  private _originalPaddingLeft?: string;

  public async onInit(): Promise<void> {

    console.log(
      '[DgaiLayout] Application Customizer inicializado'
    );

    this._sp = spfi().using(
      SPFx(this.context)
    );

    this._layoutService =
      new DgaiLayoutService(this._sp);

    try {

      this._navigationItems =
        await this._layoutService.getNavigationItems();

      console.log(
        '[DgaiLayout] Navegación:',
        this._navigationItems
      );

    } catch (error) {

      console.error(
        '[DgaiLayout] Error cargando navegación:',
        error
      );
    }

    this.context.placeholderProvider.changedEvent.add(
      this,
      this._renderPlaceholders
    );

    window.addEventListener(
      'resize',
      this._updateLayoutDimensions
    );

    this._renderPlaceholders();

    return Promise.resolve();
  }

  private _renderPlaceholders = (): void => {

    if (this._topPlaceholder) {
      return;
    }

    this._topPlaceholder =
      this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        {
          onDispose: this._onDispose
        }
      );

    if (
      !this._topPlaceholder ||
      !this._topPlaceholder.domElement
    ) {
      console.warn(
        '[DgaiLayout] Placeholder Top no disponible.'
      );

      return;
    }

    this._topPlaceholder.domElement.innerHTML = `
  <div
    id="dgai-custom-header"
    class="${styles.dgaiHeader}"
  >
        <div class="${styles.brand}">
          <span class="${styles.university}">
            Universidad de Antofagasta
          </span>

          <span class="${styles.title}">
            Dirección de Gestión y Análisis Institucional
          </span>
        </div>

        <div class="${styles.right}">
          <span class="${styles.siteName}">
            Portal DGAI
          </span>
        </div>

      </div>
    `;

    this._hideSharePointSuiteNav();
    this._hideSharePointAppBar();
    this._hideSharePointSiteHeader();
    this._hideNativeSiteNavigation();

    this._createSidebar();
    this._adjustSharePointContent();

    window.requestAnimationFrame(() => {
      this._updateLayoutDimensions();
    });
  };

  private _hideNativeSiteNavigation(): void {

    const nativeNavigation =
      document.querySelector(
        "nav[role='navigation']"
      ) as HTMLElement | null;

    if (nativeNavigation) {
      nativeNavigation.style.display =
        'none';
    }
  }

  private _createSidebar(): void {

    if (
      document.getElementById(
        'dgai-custom-sidebar'
      )
    ) {
      return;
    }

    const sidebar =
      document.createElement('aside');

    sidebar.id =
      'dgai-custom-sidebar';

    sidebar.className =
      styles.dgaiSidebar;

    const currentPath =
      window.location.pathname.toLowerCase();

    const navigationHtml =
      this._navigationItems
        .sort((a, b) => a.order - b.order)
        .map((item) => {

          const itemUrl = item.url.toLowerCase();

          const isActive =
            itemUrl !== '#' &&
            (
              currentPath === itemUrl ||
              currentPath.endsWith(itemUrl)
            );

          const cssClass = isActive
            ? `${styles.navLink} ${styles.navLinkActive}`
            : styles.navLink;

          return `
        <div class="${styles.navItemWrapper}">
          <a
            class="${cssClass}"
            href="${item.url}"
            title="${item.title}"
          >
            ${item.title}
          </a>

          ${item.separatorAfter
              ? `<div class="${styles.navSeparator}"></div>`
              : ''
            }
        </div>
      `;
        })
        .join('');

    sidebar.innerHTML = `
    <div class="${styles.sidebarTitle}">
      Navegación
    </div>

    ${navigationHtml}
  `;

    document.body.appendChild(
      sidebar
    );
  }

  private _adjustSharePointContent(): void {

    const content = document.querySelector(
      '[data-automation-id="contentScrollRegion"]'
    ) as HTMLElement | null;

    if (!content) {
      console.warn(
        '[DgaiLayout] No se encontró el área principal de contenido.'
      );
      return;
    }

    this._contentElement = content;

    this._originalMarginLeft =
      content.style.marginLeft;

    this._originalWidth =
      content.style.width;

    this._originalPaddingLeft =
      content.style.paddingLeft;

    /*
     * Ya no desplazamos el contenido.
     * SharePoint está ocupando correctamente
     * el espacio disponible después del sidebar.
     */
    content.style.marginLeft = '0';
    content.style.width = '100%';
    content.style.paddingLeft = '24px';
    content.style.boxSizing = 'border-box';
    content.style.paddingTop = '-40px';

    console.log(
      '[DgaiLayout] Área de contenido ajustada sin margen lateral adicional.'
    );
  }

  private _updateResponsiveLayout = (): void => {

    const width = window.innerWidth;

    let sidebarWidth = 220;

    if (width <= 640) {
      sidebarWidth = 0;
    } else if (width <= 900) {
      sidebarWidth = 64;
    } else if (width <= 1200) {
      sidebarWidth = 180;
    }

    document.documentElement.style.setProperty(
      '--dgai-sidebar-width',
      `${sidebarWidth}px`
    );
  };

  private _hideSharePointSiteHeader(): void {

    const hideSiteHeader = (): boolean => {

      const siteHeader = document.getElementById(
        'spSiteHeader'
      ) as HTMLElement | null;

      if (!siteHeader) {
        return false;
      }

      if (!this._siteHeaderElement) {
        this._siteHeaderElement = siteHeader;

        this._originalSiteHeaderDisplay =
          siteHeader.style.display;
      }

      siteHeader.style.setProperty(
        'display',
        'none',
        'important'
      );

      return true;
    };

    // Primer intento
    hideSiteHeader();

    // SharePoint puede volver a crear el header
    this._siteHeaderObserver =
      new MutationObserver(() => {
        hideSiteHeader();
      });

    this._siteHeaderObserver.observe(
      document.body,
      {
        childList: true,
        subtree: true
      }
    );
  }

  private _hideSharePointAppBar(): void {

    const hideAppBar = (): boolean => {

      const appBar = document.getElementById(
        'sp-appBar'
      ) as HTMLElement | null;

      if (!appBar) {
        return false;
      }

      if (!this._appBarElement) {
        this._appBarElement = appBar;

        this._originalAppBarDisplay =
          appBar.style.display;
      }

      appBar.style.setProperty(
        'display',
        'none',
        'important'
      );

      console.log(
        '[DgaiLayout] Barra lateral de SharePoint ocultada.'
      );

      return true;
    };

    // Intentar inmediatamente
    hideAppBar();

    // SharePoint puede crear/recrear el elemento posteriormente
    this._appBarObserver =
      new MutationObserver(() => {
        hideAppBar();
      });

    this._appBarObserver.observe(
      document.body,
      {
        childList: true,
        subtree: true
      }
    );
  }

  private _updateLayoutDimensions = (): void => {

    const viewportWidth = window.innerWidth;

    let sidebarWidth = 260;

    if (viewportWidth <= 1200) {
      sidebarWidth = 220;
    }

    if (viewportWidth <= 900) {
      sidebarWidth = 72;
    }

    if (viewportWidth <= 640) {
      sidebarWidth = 0;
    }

    document.documentElement.style.setProperty(
      '--dgai-sidebar-width',
      `${sidebarWidth}px`
    );

    const header = document.getElementById(
      'dgai-custom-header'
    ) as HTMLElement | null;

    if (!header) {
      console.warn(
        '[DgaiLayout] No se encontró el header DGAI.'
      );
      return;
    }

    const headerRect =
      header.getBoundingClientRect();

    document.documentElement.style.setProperty(
      '--dgai-header-bottom',
      `${headerRect.bottom}px`
    );

    console.log(
      '[DgaiLayout] Header bottom:',
      headerRect.bottom
    );
  };

  private _hideSharePointSuiteNav(): void {

    const hideSuiteNav = (): boolean => {

      const suiteNav = document.getElementById(
        'SuiteNavWrapper'
      ) as HTMLElement | null;

      if (!suiteNav) {
        return false;
      }

      if (!this._suiteNavElement) {
        this._suiteNavElement = suiteNav;

        this._originalSuiteNavDisplay =
          suiteNav.style.display;
      }

      suiteNav.style.setProperty(
        'display',
        'none',
        'important'
      );

      return true;
    };

    // Intentar inmediatamente
    hideSuiteNav();

    // SharePoint puede volver a crear el elemento
    this._suiteNavObserver =
      new MutationObserver(() => {
        hideSuiteNav();
      });

    this._suiteNavObserver.observe(
      document.body,
      {
        childList: true,
        subtree: true
      }
    );
  }

  private _onDispose(): void {

    const sidebar = document.getElementById(
      'dgai-custom-sidebar'
    );

    sidebar?.remove();

    window.removeEventListener(
      'resize',
      this._updateResponsiveLayout
    );

    document.documentElement.style.removeProperty(
      '--dgai-sidebar-width'
    );

    if (this._contentElement) {

      this._contentElement.style.marginLeft =
        this._originalMarginLeft || '';

      this._contentElement.style.width =
        this._originalWidth || '';

      this._contentElement.style.paddingLeft =
        this._originalPaddingLeft || '';

      this._contentElement.style.boxSizing = '';
    }
    if (this._siteHeaderElement) {

      this._siteHeaderElement.style.marginLeft =
        this._originalSiteHeaderMarginLeft || '';

      this._siteHeaderElement.style.width =
        this._originalSiteHeaderWidth || '';

      this._siteHeaderElement.style.boxSizing = '';
    }
    if (this._siteHeaderElement) {
      this._siteHeaderElement.style.display =
        this._originalSiteHeaderDisplay || '';
    } if (this._appBarObserver) {
      this._appBarObserver.disconnect();
      this._appBarObserver = undefined;
    }
    if (this._appBarElement) {

      this._appBarElement.style.removeProperty(
        'display'
      );

      if (this._originalAppBarDisplay) {
        this._appBarElement.style.display =
          this._originalAppBarDisplay;
      }
    }

    if (this._suiteNavObserver) {
      this._suiteNavObserver.disconnect();
      this._suiteNavObserver = undefined;
    }
    if (this._suiteNavElement) {

      this._suiteNavElement.style.removeProperty(
        'display'
      );

      if (this._originalSuiteNavDisplay) {
        this._suiteNavElement.style.display =
          this._originalSuiteNavDisplay;
      }
    }

    window.removeEventListener(
      'resize',
      this._updateLayoutDimensions
    );

    document.documentElement.style.removeProperty(
      '--dgai-sidebar-width'
    );

    document.documentElement.style.removeProperty(
      '--dgai-header-bottom'
    );

    console.log(
      '[DgaiLayout] Extension eliminada.'
    );
  }
}