import type { DgaiPortalService } from '../services/DgaiPortalService';
import type { DgaiHomeService } from '../services/DgaiHomeService';

export interface IDgaiPortalProps {
  portalTitle: string;
  portalDescription: string;
  showQuickLinks: boolean;
  homeService: DgaiHomeService;
  portalService: DgaiPortalService;
}