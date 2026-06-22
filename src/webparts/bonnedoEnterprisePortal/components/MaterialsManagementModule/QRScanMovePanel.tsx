import * as React from 'react';
import {
  DefaultButton,
  Dropdown,
  getTheme,
  Icon,
  IDropdownOption,
  Image,
  ImageFit,
  MessageBar,
  MessageBarType,
  mergeStyleSets,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  Stack,
  Text,
  TextField,
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { SharePointService } from '../../services/SharePointService';
import { SHAREPOINT_LISTS } from '../../constants/SharePointListNames';
import { QRCodeService } from '../../services/QRCodeService';
import { IMaterialMasterRecord } from '../../models/DataModels';

export type QRScanSource = 'qr-tab' | 'inventory' | 'movements';

export type InventoryMovementType = 'GRN' | 'Transfer Out' | 'Issue' | 'Return' | 'Verify' | 'Scrap';

export interface IQRScanRequestContext {
  source?: QRScanSource;
  movementType?: InventoryMovementType;
  location?: string;
}

export interface IQRScanMovePanelProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  userDisplayName: string;
  context?: IQRScanRequestContext;
  onScanComplete?: () => void;
}

interface IWarehouseOption {
  code: string;
  name: string;
  status: string;
}

interface IBarcodeDetector {
  detect(source: HTMLVideoElement): Promise<Array<{ rawValue: string }>>;
}

interface IBarcodeDetectorConstructor {
  new (options?: { formats?: string[] }): IBarcodeDetector;
}

interface IInventoryAdjustment {
  location: string;
  delta: number;
}

interface IWorkflowStep {
  label: string;
  key: string;
  icon: string;
}

const MOVEMENT_TYPES: Array<{ key: InventoryMovementType; text: string; icon: string; label: string }> = [
  { key: 'GRN', text: 'Receive', icon: 'Inbox', label: 'Receive' },
  { key: 'Transfer Out', text: 'Transfer', icon: 'GlobalTransferBus', label: 'Transfer' },
  { key: 'Issue', text: 'Issue', icon: 'ArrowUpRight8', label: 'Issue' },
  { key: 'Return', text: 'Return', icon: 'ReturnToSession', label: 'Return' },
  { key: 'Verify', text: 'Verify', icon: 'Accept', label: 'Verify' },
  { key: 'Scrap', text: 'Scrap', icon: 'Blocked2', label: 'Scrap' },
];

const DEFAULT_CONTEXT: Required<Pick<IQRScanRequestContext, 'movementType'>> = {
  movementType: 'GRN',
};

const WORKFLOW_STEPS: IWorkflowStep[] = [
  { label: 'Scan QR', key: 'scan', icon: 'QRCode' },
  { label: 'View Item', key: 'view', icon: 'Info' },
  { label: 'Action', key: 'action', icon: 'ActionCenter' },
  { label: 'Confirm', key: 'confirm', icon: 'Accept' },
  { label: 'Updated', key: 'updated', icon: 'Refresh' },
];

const QR_SCAN_STYLES = mergeStyleSets({
  pageHead: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: 20,
  },
  pageTitle: {
    fontSize: 18,
    fontWeight: 600,
  },
  pageSubtitle: {
    fontSize: 12,
    color: '#7F8C9B',
    marginTop: 2,
  },
  twoColumn: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, minmax(0, 1fr))',
    gap: 16,
  },
  scanBox: {
    background: 'linear-gradient(135deg, #141920 0%, #1E2532 100%)',
    borderRadius: 8,
    padding: 48,
    textAlign: 'center',
    minHeight: 300,
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    position: 'relative',
    overflow: 'hidden',
  },
  scanVideo: {
    width: '100%',
    maxHeight: 320,
    objectFit: 'cover',
    borderRadius: 8,
  },
  scanFrame: {
    width: 180,
    height: 180,
    border: '2px solid rgba(0,184,148,0.4)',
    borderRadius: 14,
    position: 'relative',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
  },
  scanLine: {
    position: 'absolute',
    top: 0,
    left: '15%',
    right: '15%',
    height: 2,
    background: '#00B894',
    boxShadow: '0 0 12px #00B894',
    animation: 'scanAnim 2s ease-in-out infinite',
  },
  scanHint: {
    color: 'rgba(255,255,255,0.45)',
    fontSize: 12,
    marginTop: 16,
  },
  workflowCard: {
    backgroundColor: '#FFFFFF',
    borderRadius: 8,
    boxShadow: '0 1px 3px rgba(0,0,0,0.04), 0 1px 2px rgba(0,0,0,0.02)',
    marginTop: 12,
  },
  workflowBody: {
    padding: '12px 16px',
  },
  workflowLabel: {
    fontSize: 12,
    fontWeight: 600,
    marginBottom: 6,
  },
  workflowRow: {
    display: 'flex',
    alignItems: 'center',
    gap: 6,
    flexWrap: 'wrap',
    margin: '10px 0',
  },
  workflowStep: {
    backgroundColor: '#FFFFFF',
    border: '1px solid #E8ECF0',
    borderRadius: 6,
    padding: '8px 12px',
    fontSize: 11,
    textAlign: 'center',
  },
  workflowStepActive: {
    borderColor: '#00B894',
    backgroundColor: 'rgba(0,184,148,0.08)',
    fontWeight: 600,
    color: '#00B894',
  },
  workflowArrow: {
    color: '#B0B8C4',
    fontSize: 11,
  },
  resultCard: {
    backgroundColor: '#FFFFFF',
    borderRadius: 8,
    border: '1px solid #E8ECF0',
    overflow: 'hidden',
    boxShadow: '0 1px 3px rgba(0,0,0,0.04), 0 1px 2px rgba(0,0,0,0.02)',
  },
  resultTop: {
    backgroundColor: '#00B894',
    color: '#FFFFFF',
    padding: '12px 16px',
    display: 'flex',
    alignItems: 'center',
    gap: 8,
    fontSize: 13,
    fontWeight: 600,
  },
  resultBody: {
    padding: 16,
  },
  materialCode: {
    fontSize: 15,
    fontWeight: 700,
    fontFamily: "'Cascadia Code', 'Fira Code', Consolas, monospace",
  },
  materialDesc: {
    fontSize: 12,
    color: '#7F8C9B',
    marginBottom: 14,
  },
  detailGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, minmax(0, 1fr))',
    gap: 10,
  },
  detailItem: {
    padding: 8,
    borderRadius: 6,
    backgroundColor: '#F5F6FA',
  },
  detailLabel: {
    fontSize: 10,
    color: '#B0B8C4',
    textTransform: 'uppercase',
    letterSpacing: '0.03em',
    fontWeight: 600,
  },
  detailValue: {
    fontSize: 13,
    fontWeight: 500,
    marginTop: 1,
    wordBreak: 'break-word',
  },
  actionTitle: {
    fontSize: 10,
    fontWeight: 600,
    margin: '14px 0 8px',
    color: '#B0B8C4',
    textTransform: 'uppercase',
    letterSpacing: '0.05em',
  },
  actionGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, minmax(0, 1fr))',
    gap: 8,
    margin: '14px 0',
  },
  actionCard: {
    border: '1px solid #E8ECF0',
    borderRadius: 8,
    padding: '14px 8px',
    textAlign: 'center',
    cursor: 'pointer',
    transition: 'all 0.15s',
    backgroundColor: '#FFFFFF',
  },
  actionCardHover: {
    borderColor: '#00B894',
    backgroundColor: 'rgba(0,184,148,0.08)',
  },
  actionCardSelected: {
    borderColor: '#00B894',
    backgroundColor: 'rgba(0,184,148,0.08)',
    boxShadow: '0 0 0 2px rgba(0,184,148,0.15)',
  },
  actionCardIcon: {
    fontSize: 18,
    color: '#00B894',
    display: 'block',
    marginBottom: 5,
  },
  actionCardText: {
    fontSize: 11,
    fontWeight: 500,
    color: '#7F8C9B',
  },
  actionControls: {
    display: 'flex',
    gap: 8,
    marginTop: 12,
    flexWrap: 'wrap',
    alignItems: 'flex-end',
  },
  compactDropdown: {
    minWidth: 160,
  },
  compactTextField: {
    maxWidth: 90,
  },
  qrPanel: {
    display: 'flex',
    alignItems: 'center',
    gap: 12,
    maxWidth: 280,
    border: '2px dashed #E8ECF0',
    borderRadius: 8,
    padding: '12px 14px',
    marginTop: 12,
  },
  qrBox: {
    width: 56,
    height: 56,
    backgroundColor: '#F5F6FA',
    border: '1px solid #E8ECF0',
    borderRadius: 4,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
  },
  qrInfo: {
    fontFamily: "'Cascadia Code', 'Fira Code', Consolas, monospace",
    fontSize: 10,
  },
  qrCode: {
    fontWeight: 700,
    fontSize: 11,
    color: '#2C3E50',
  },
  qrDesc: {
    color: '#7F8C9B',
  },
  qrLoc: {
    color: '#B0B8C4',
    fontSize: 9,
  },
  confirmBox: {
    backgroundColor: '#FFFFFF',
    borderRadius: 8,
    border: '1px solid #E8ECF0',
    overflow: 'hidden',
    boxShadow: '0 1px 3px rgba(0,0,0,0.04), 0 1px 2px rgba(0,0,0,0.02)',
    textAlign: 'center',
    padding: 32,
  },
  confirmIcon: {
    width: 52,
    height: 52,
    borderRadius: '50%',
    backgroundColor: 'rgba(0,184,148,0.08)',
    color: '#00B894',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    margin: '0 auto 12px',
  },
  emptyCard: {
    backgroundColor: '#FFFFFF',
    borderRadius: 8,
    border: '1px solid #E8ECF0',
    boxShadow: '0 1px 3px rgba(0,0,0,0.04), 0 1px 2px rgba(0,0,0,0.02)',
    padding: 32,
    textAlign: 'center',
  },
  tag: {
    display: 'inline-block',
    padding: '3px 10px',
    borderRadius: 20,
    fontSize: 10,
    fontWeight: 600,
  },
  tagGreen: {
    backgroundColor: 'rgba(0,184,148,0.08)',
    color: '#00B894',
  },
  tagBlue: {
    backgroundColor: 'rgba(74,144,217,0.08)',
    color: '#4A90D9',
  },
  tagOrange: {
    backgroundColor: 'rgba(243,156,18,0.08)',
    color: '#F39C12',
  },
  tagRed: {
    backgroundColor: 'rgba(231,76,60,0.08)',
    color: '#E74C3C',
  },
  tagPurple: {
    backgroundColor: 'rgba(155,89,182,0.08)',
    color: '#9B59B6',
  },
  loadingCard: {
    backgroundColor: '#FFFFFF',
    borderRadius: 8,
    border: '1px solid #E8ECF0',
    boxShadow: '0 1px 3px rgba(0,0,0,0.04), 0 1px 2px rgba(0,0,0,0.02)',
    padding: 32,
  },
});

function parseMaterialCodeFromScan(rawValue: string): string {
  const trimmed = rawValue.trim();

  if (!trimmed) {
    return '';
  }

  let parsed: Record<string, unknown> | undefined;

  try {
    parsed = JSON.parse(trimmed) as Record<string, unknown>;
  } catch {
    parsed = undefined;
  }

  if (parsed) {
    const materialCode =
      parsed.materialCode ||
      parsed.Material_Code ||
      parsed.material ||
      parsed.code ||
      parsed.itemCode ||
      parsed.itemId ||
      parsed.qrcodeurl ||
      parsed.QRCodeURL;

    if (typeof materialCode === 'string' && materialCode.trim()) {
      return materialCode.trim();
    }
  }

  let scannedUrl: URL | undefined;

  try {
    scannedUrl = new URL(trimmed);
  } catch {
    scannedUrl = undefined;
  }

  if (scannedUrl) {
    const encodedData = scannedUrl.searchParams.get('data');

    if (encodedData) {
      return parseMaterialCodeFromScan(decodeURIComponent(encodedData));
    }
  }

  const patterns = [
    /^bonnedo:material:(.+)$/i,
    /^bonnedo-material:(.+)$/i,
    /^material:(.+)$/i,
    /^mat:(.+)$/i,
  ];

  for (const pattern of patterns) {
    const match = trimmed.match(pattern);

    if (match?.[1]) {
      return match[1].trim();
    }
  }

  return trimmed;
}

function getMaterialQRCodeUrl(material: IMaterialMasterRecord): string {
  return material.QRCodeURL || material.qrcodeurl || material.MaterialQRCode || material.Material_QR_Code || '';
}

function findMaterialByCodeOrQr(
  materials: IMaterialMasterRecord[],
  materialCode: string
): IMaterialMasterRecord | undefined {
  const normalized = materialCode.trim().toLowerCase();

  if (!normalized) {
    return undefined;
  }

  return materials.find((material) => {
    const code = (material.Material_Code || '').toLowerCase();
    const title = (material.Title || '').toLowerCase();
    const qrCodeUrl = getMaterialQRCodeUrl(material).toLowerCase();

    return code === normalized || title === normalized || qrCodeUrl === normalized;
  });
}

function getMovementAdjustments(
  movementType: InventoryMovementType,
  fromLocation: string,
  toLocation: string,
  qty: number
): IInventoryAdjustment[] {
  switch (movementType) {
    case 'GRN':
      return [{ location: toLocation, delta: qty }];
    case 'Transfer Out':
      return [
        { location: fromLocation, delta: -qty },
        { location: toLocation, delta: qty },
      ];
    case 'Issue':
      return [{ location: fromLocation, delta: -qty }];
    case 'Return':
      return [{ location: toLocation, delta: qty }];
    case 'Verify':
      return [];
    case 'Scrap':
      return [{ location: fromLocation, delta: -qty }];
    default:
      return [];
  }
}

function createBarcodeDetector(): IBarcodeDetector | undefined {
  const Detector = (window as unknown as { BarcodeDetector?: IBarcodeDetectorConstructor }).BarcodeDetector;

  if (!Detector) {
    return undefined;
  }

  try {
    return new Detector({ formats: ['qr_code'] });
  } catch {
    return undefined;
  }
}

function getCategoryTagClass(category: string): keyof typeof QR_SCAN_STYLES {
  const normalized = (category || '').toLowerCase();

  if (normalized.includes('valve')) {
    return 'tagRed';
  }

  if (normalized.includes('pipe')) {
    return 'tagBlue';
  }

  if (normalized.includes('flange')) {
    return 'tagOrange';
  }

  if (normalized.includes('reducer')) {
    return 'tagPurple';
  }

  return 'tagGreen';
}

function getMovementTagClass(movementType: InventoryMovementType): keyof typeof QR_SCAN_STYLES {
  switch (movementType) {
    case 'GRN':
      return 'tagGreen';
    case 'Transfer Out':
      return 'tagBlue';
    case 'Issue':
      return 'tagOrange';
    case 'Return':
      return 'tagPurple';
    case 'Scrap':
      return 'tagRed';
    default:
      return 'tagBlue';
  }
}

const QRScanMovePanel: React.FC<IQRScanMovePanelProps> = ({
  spHttpClient,
  pageContext,
  userDisplayName,
  context,
  onScanComplete,
}) => {
  const theme = getTheme();
  const videoRef = React.useRef<HTMLVideoElement | null>(null);
  const streamRef = React.useRef<MediaStream | null>(null);
  const detectionTimerRef = React.useRef<number | null>(null);

  const [materialOptions, setMaterialOptions] = React.useState<IMaterialMasterRecord[]>([]);
  const [warehouseOptions, setWarehouseOptions] = React.useState<IWarehouseOption[]>([]);
  const [referenceLoading, setReferenceLoading] = React.useState(false);
  const [saving, setSaving] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
  const [scanStatus, setScanStatus] = React.useState('');
  const [cameraActive, setCameraActive] = React.useState(false);
  const [materialCode, setMaterialCode] = React.useState('');
  const [quantity, setQuantity] = React.useState('1');
  const [movementType, setMovementType] = React.useState<InventoryMovementType>(
    context?.movementType || DEFAULT_CONTEXT.movementType
  );
  const [fromLocation, setFromLocation] = React.useState('');
  const [toLocation, setToLocation] = React.useState('');
  const [projectCode, setProjectCode] = React.useState('');
  const [note, setNote] = React.useState('');
  const [confirmedMovement, setConfirmedMovement] = React.useState<{
    movementType: InventoryMovementType;
    materialCode: string;
    quantity: number;
    fromLocation?: string;
    toLocation?: string;
  } | null>(null);

  const sharePointService = React.useMemo(
    () => new SharePointService(spHttpClient, pageContext),
    [spHttpClient, pageContext]
  );

  const activeWarehouse = React.useMemo(
    () => warehouseOptions.find((warehouse) => warehouse.status === 'Active') || warehouseOptions[0],
    [warehouseOptions]
  );

  const selectedMaterial = React.useMemo(
    () => findMaterialByCodeOrQr(materialOptions, materialCode),
    [materialCode, materialOptions]
  );

  const qrCodeUrl = React.useMemo(() => {
    const code = selectedMaterial?.Material_Code || materialCode;

    if (!code) {
      return '';
    }

    return selectedMaterial ? getMaterialQRCodeUrl(selectedMaterial) || QRCodeService.generateQRCodeUrl(code, 160) : QRCodeService.generateQRCodeUrl(code, 160);
  }, [materialCode, selectedMaterial]);

  const currentWorkflowStep = React.useMemo((): string => {
    if (confirmedMovement) {
      return 'updated';
    }

    if (saving) {
      return 'confirm';
    }

    if (!selectedMaterial) {
      return 'scan';
    }

    if (!movementType) {
      return 'view';
    }

    return 'action';
  }, [confirmedMovement, movementType, saving, selectedMaterial]);

  const loadReferenceData = React.useCallback(async (): Promise<void> => {
    setReferenceLoading(true);
    setError(null);

    try {
      const [materials, warehouses] = await Promise.all([
        sharePointService.getMaterialMasterRecords(),
        sharePointService.getWarehouses().catch(() => []),
      ]);

      setMaterialOptions(materials);
      setWarehouseOptions(
        warehouses.map((warehouse) => ({
          code: warehouse.WarehouseCode || warehouse.Title || '',
          name: warehouse.Warehouse_Name || warehouse.Title || '',
          status: warehouse.Status || 'Active',
        }))
      );
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to load scan reference data');
    } finally {
      setReferenceLoading(false);
    }
  }, [sharePointService]);

  React.useEffect(() => {
    loadReferenceData().catch(() => undefined);
  }, [loadReferenceData]);

  React.useEffect(() => {
    if (cameraActive && videoRef.current && streamRef.current) {
      videoRef.current.srcObject = streamRef.current;
    }
  }, [cameraActive]);

  React.useEffect(() => {
    if (context?.movementType) {
      setMovementType(context.movementType);
    }
  }, [context?.movementType]);

  React.useEffect(() => {
    if (!activeWarehouse) {
      return;
    }

    if (context?.location && (!fromLocation || !toLocation)) {
      setFromLocation((current) => current || context.location || '');
      setToLocation((current) => current || context.location || '');
      return;
    }

    if (movementType === 'GRN' || movementType === 'Return') {
      setToLocation((current) => current || activeWarehouse.code);
    }

    if (movementType === 'Issue' || movementType === 'Transfer Out' || movementType === 'Scrap') {
      setFromLocation((current) => current || activeWarehouse.code);
    }
  }, [activeWarehouse, context?.location, movementType, fromLocation, toLocation]);

  const handleScanResult = React.useCallback(
    (rawValue: string): void => {
      if (referenceLoading) {
        setScanStatus('Loading material master records...');
        return;
      }

      const parsedMaterialCode = parseMaterialCodeFromScan(rawValue);

      if (!parsedMaterialCode) {
        setScanStatus('No material code found in the scanned QR code.');
        return;
      }

      const material = findMaterialByCodeOrQr(materialOptions, parsedMaterialCode);

      if (!material) {
        setError(`Material ${parsedMaterialCode} was not found in the material master.`);
        setScanStatus('Scanned code is not linked to a material master record.');
        return;
      }

      setMaterialCode(material.Material_Code || parsedMaterialCode);
      setConfirmedMovement(null);
      setScanStatus(`Scanned material: ${material.Material_Code}`);
    },
    [materialOptions, referenceLoading]
  );

  React.useEffect(() => {
    if (!cameraActive) {
      return undefined;
    }

    const detector = createBarcodeDetector();

    if (!detector || !videoRef.current) {
      return undefined;
    }

    let detectionErrorShown = false;

    const detectFrame = async (): Promise<void> => {
      if (!videoRef.current || !streamRef.current) {
        return;
      }

      try {
        const barcodes = await detector.detect(videoRef.current);

        if (barcodes.length > 0 && barcodes[0].rawValue) {
          handleScanResult(barcodes[0].rawValue);
        }
      } catch {
        if (!detectionErrorShown) {
          setScanStatus('Waiting for a readable camera frame.');
          detectionErrorShown = true;
        }
      }
    };

    detectionTimerRef.current = window.setInterval(() => {
      detectFrame().catch(() => undefined);
    }, 500);

    return (): void => {
      if (detectionTimerRef.current !== null) {
        window.clearInterval(detectionTimerRef.current);
        detectionTimerRef.current = null;
      }
    };
  }, [cameraActive, handleScanResult]);

  React.useEffect(() => {
    return (): void => {
      if (detectionTimerRef.current !== null) {
        window.clearInterval(detectionTimerRef.current);
        detectionTimerRef.current = null;
      }

      if (streamRef.current) {
        streamRef.current.getTracks().forEach((track) => track.stop());
        streamRef.current = null;
      }

      if (videoRef.current) {
        videoRef.current.srcObject = null;
      }
    };
  }, []);

  const stopCamera = React.useCallback((): void => {
    if (detectionTimerRef.current !== null) {
      window.clearInterval(detectionTimerRef.current);
      detectionTimerRef.current = null;
    }

    if (streamRef.current) {
      streamRef.current.getTracks().forEach((track) => track.stop());
      streamRef.current = null;
    }

    setCameraActive(false);

    if (videoRef.current) {
      videoRef.current.srcObject = null;
    }

    setScanStatus('Camera stopped.');
  }, []);

  const startCamera = React.useCallback(async (): Promise<void> => {
    stopCamera();
    setError(null);
    setScanStatus('Requesting camera access...');

    if (!window.isSecureContext) {
      setError('Camera scanning requires a secure HTTPS connection.');
      return;
    }

    if (!navigator.mediaDevices?.getUserMedia) {
      setError('Camera scanning is not available in this browser.');
      return;
    }

    try {
      const stream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: 'environment' },
        audio: false,
      });

      streamRef.current = stream;
      setCameraActive(true);

      if (videoRef.current) {
        videoRef.current.srcObject = stream;
      }

      const detector = createBarcodeDetector();

      if (detector) {
        setScanStatus('Camera active. Point the camera at a material QR code.');
      } else {
        setScanStatus('Camera active. Manual entry is available because QR detection is not supported in this browser.');
      }
    } catch (err) {
      setScanStatus('Camera unavailable.');
      setError(err instanceof Error ? err.message : 'Failed to start camera');
    }
  }, [stopCamera]);

  const handleManualScan = React.useCallback((): void => {
    const parsedMaterialCode = parseMaterialCodeFromScan(materialCode);

    if (!parsedMaterialCode) {
      setError('Enter or scan a material code.');
      return;
    }

    const material = findMaterialByCodeOrQr(materialOptions, parsedMaterialCode);

    if (!material) {
      setError(`Material ${parsedMaterialCode} was not found in the material master.`);
      return;
    }

    setMaterialCode(material.Material_Code || parsedMaterialCode);
    setConfirmedMovement(null);
    setError(null);
    setScanStatus(`Material loaded: ${material.Material_Code}`);
  }, [materialCode, materialOptions]);

  const resetScan = React.useCallback((): void => {
    setMaterialCode('');
    setConfirmedMovement(null);
    setError(null);
    setScanStatus('Scan or enter the next material code.');
  }, []);

  const updateInventoryForMovement = React.useCallback(
    async (
      movement: InventoryMovementType,
      code: string,
      qty: number,
      from: string,
      to: string,
      project: string,
      timestamp: string
    ): Promise<void> => {
      if (movement === 'Transfer Out' && from && to && from === to) {
        throw new Error('Choose different locations for transfer movements.');
      }

      const adjustments = getMovementAdjustments(movement, from, to, qty);

      if (adjustments.some((adjustment) => !adjustment.location)) {
        throw new Error('Select a warehouse location for this movement.');
      }

      const inventoryRecords = await sharePointService.getInventoryRecords();

      for (const adjustment of adjustments) {
        const existingRecord = inventoryRecords.find((record) => {
          const matchesMaterial = record.Material_Code === code;
          const matchesLocation = record.Location === adjustment.location;
          const matchesProject = project ? record.Project_Code === project : !record.Project_Code;
          return matchesMaterial && matchesLocation && matchesProject;
        });

        if (!existingRecord) {
          if (adjustment.delta < 0) {
            throw new Error(`No inventory found for ${code} at ${adjustment.location}.`);
          }

          await sharePointService.createListItem(SHAREPOINT_LISTS.INVENTORY_REGISTER, {
            Title: `${code} - ${adjustment.location}`,
            Material_Code: code,
            Location: adjustment.location,
            Project_Code: project || null,
            Qty_On_Hand: adjustment.delta,
            QtyReserved: 0,
            Condition: 'Good',
            Status: 'Available',
            DateReceived: timestamp,
            Last_Movement_Date: timestamp,
          });
          continue;
        }

        const currentQty = Number(existingRecord.Qty_On_Hand || 0);
        const nextQty = currentQty + adjustment.delta;

        if (nextQty < 0) {
          throw new Error(`Insufficient stock for ${code} at ${adjustment.location}.`);
        }

        await sharePointService.updateListItem(SHAREPOINT_LISTS.INVENTORY_REGISTER, existingRecord.ID, {
          Qty_On_Hand: nextQty,
          Last_Movement_Date: timestamp,
          Status: nextQty === 0 ? 'Available' : existingRecord.Status || 'Available',
        });
      }
    },
    [sharePointService]
  );

  const handleSaveMovement = React.useCallback(async (): Promise<void> => {
    if (saving || referenceLoading) {
      return;
    }

    setError(null);

    const parsedMaterialCode = parseMaterialCodeFromScan(materialCode);

    if (!parsedMaterialCode) {
      setError('Scan or enter a material code before confirming the movement.');
      return;
    }

    const parsedQuantity = Number(quantity);

    if (!Number.isFinite(parsedQuantity) || parsedQuantity <= 0) {
      setError('Enter a quantity greater than zero.');
      return;
    }

    const material = selectedMaterial || findMaterialByCodeOrQr(materialOptions, parsedMaterialCode);

    if (!material) {
      setError(`Material ${parsedMaterialCode} was not found in the material master.`);
      return;
    }

    if (material.Active === false) {
      setError(`Material ${parsedMaterialCode} is inactive. Reactivate it before recording movements.`);
      return;
    }

    if (movementType === 'Verify' && !note.trim()) {
      setError('Add a verification note before recording a Verify movement.');
      return;
    }

    const timestamp = new Date().toISOString();
    const movementNote = [
      note,
      `Scanned by ${userDisplayName} at ${new Date().toLocaleString('en-GB')}`,
    ]
      .filter(Boolean)
      .join('\n');

    setSaving(true);

    try {
      await sharePointService.createListItem(SHAREPOINT_LISTS.INVENTORY_MOVEMENTS_REGISTER, {
        Title: `${movementType} - ${parsedMaterialCode}`,
        Movement_Type: movementType,
        Material_Code: parsedMaterialCode,
        Qty: parsedQuantity,
        From_Location: fromLocation || null,
        To_Location: toLocation || null,
        Project_Code: projectCode || null,
        Note: movementNote || null,
      });

      await updateInventoryForMovement(
        movementType,
        parsedMaterialCode,
        parsedQuantity,
        fromLocation,
        toLocation,
        projectCode,
        timestamp
      );

      setConfirmedMovement({
        movementType,
        materialCode: parsedMaterialCode,
        quantity: parsedQuantity,
        fromLocation,
        toLocation,
      });
      setScanStatus('Movement recorded. Lists updated.');
      onScanComplete?.();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to save QR scan movement');
    } finally {
      setSaving(false);
    }
  }, [
    saving,
    referenceLoading,
    materialCode,
    quantity,
    selectedMaterial,
    materialOptions,
    note,
    userDisplayName,
    movementType,
    fromLocation,
    toLocation,
    projectCode,
    sharePointService,
    updateInventoryForMovement,
    onScanComplete,
  ]);

  const warehouseOptionsForDropdown: IDropdownOption[] = warehouseOptions.map((warehouse) => ({
    key: warehouse.code,
    text: warehouse.code ? `${warehouse.code} — ${warehouse.name}` : warehouse.name,
  }));

  const movementDetail = React.useMemo(() => {
    if (!confirmedMovement) {
      return '';
    }

    const locations = [confirmedMovement.fromLocation, confirmedMovement.toLocation].filter(Boolean).join(' → ');
    return `${confirmedMovement.materialCode} — ${confirmedMovement.movementType}${locations ? ` (${locations}, ${confirmedMovement.quantity})` : ` (${confirmedMovement.quantity})`}`;
  }, [confirmedMovement]);

  const canSave = Boolean(selectedMaterial) && Boolean(movementType) && !referenceLoading && !saving;

  return (
    <>
      <style>
        {`
          @keyframes scanAnim {
            0%, 100% { top: 0; }
            50% { top: 100%; }
          }

          @media (max-width: 900px) {
            .${QR_SCAN_STYLES.twoColumn} {
              grid-template-columns: 1fr;
            }
          }
        `}
      </style>

      <Stack tokens={{ childrenGap: 16 }}>
        {error && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline onDismiss={() => setError(null)}>
            {error}
          </MessageBar>
        )}

        <div className={QR_SCAN_STYLES.pageHead}>
          <div>
            <div className={QR_SCAN_STYLES.pageTitle}>QR Scan &amp; Move</div>
            <div className={QR_SCAN_STYLES.pageSubtitle}>
              {scanStatus || 'Scan material QR code to record movements'}
            </div>
          </div>
        </div>

        {referenceLoading ? (
          <Stack horizontalAlign="center" verticalAlign="center" className={QR_SCAN_STYLES.loadingCard}>
            <Spinner size={SpinnerSize.medium} label="Loading materials and warehouses..." />
          </Stack>
        ) : (
          <div className={QR_SCAN_STYLES.twoColumn}>
            <div>
              <div className={QR_SCAN_STYLES.scanBox}>
                {cameraActive ? (
                  <>
                    <video
                      ref={videoRef}
                      autoPlay
                      muted
                      playsInline
                      className={QR_SCAN_STYLES.scanVideo}
                    />
                    <div className={QR_SCAN_STYLES.scanFrame}>
                      <div className={QR_SCAN_STYLES.scanLine} />
                    </div>
                  </>
                ) : (
                  <>
                    <div className={QR_SCAN_STYLES.scanFrame}>
                      <Icon
                        iconName="QRCode"
                        style={{ fontSize: 40, color: 'rgba(255,255,255,0.12)' }}
                      />
                    </div>
                    <div className={QR_SCAN_STYLES.scanHint}>Position QR code within frame</div>
                    <PrimaryButton
                      iconProps={{ iconName: 'Camera' }}
                      onClick={() => startCamera().catch(() => undefined)}
                      disabled={saving}
                      styles={{ root: { marginTop: 20 } }}
                    >
                      Start Camera
                    </PrimaryButton>
                  </>
                )}

                {cameraActive && (
                  <DefaultButton
                    text="Stop camera"
                    iconProps={{ iconName: 'StopSolid' }}
                    onClick={stopCamera}
                    disabled={saving}
                    styles={{ root: { marginTop: 16 } }}
                  />
                )}
              </div>

              <div className={QR_SCAN_STYLES.workflowCard}>
                <div className={QR_SCAN_STYLES.workflowBody}>
                  <div className={QR_SCAN_STYLES.workflowLabel}>Workflow</div>
                  <div className={QR_SCAN_STYLES.workflowRow}>
                    {WORKFLOW_STEPS.map((step, index) => (
                      <React.Fragment key={step.key}>
                        <div
                          className={`${QR_SCAN_STYLES.workflowStep} ${
                            currentWorkflowStep === step.key ? QR_SCAN_STYLES.workflowStepActive : ''
                          }`}
                        >
                          <Icon iconName={step.icon} /> {step.label}
                        </div>
                        {index < WORKFLOW_STEPS.length - 1 && (
                          <div className={QR_SCAN_STYLES.workflowArrow}>→</div>
                        )}
                      </React.Fragment>
                    ))}
                  </div>
                </div>
              </div>

              <div className={QR_SCAN_STYLES.workflowCard}>
                <div className={QR_SCAN_STYLES.workflowBody}>
                  <TextField
                    label="Scanned material code"
                    value={materialCode}
                    onChange={(_, value) => {
                      setMaterialCode(value || '');
                      setConfirmedMovement(null);
                    }}
                    placeholder="Scan a QR code or paste the material code"
                    disabled={saving}
                    styles={{ root: { marginBottom: 8 } }}
                  />
                  <DefaultButton
                    text="Load material code"
                    onClick={handleManualScan}
                    disabled={saving || referenceLoading || !materialCode}
                  />
                </div>
              </div>
            </div>

            <div>
              {confirmedMovement ? (
                <div className={QR_SCAN_STYLES.confirmBox}>
                  <div className={QR_SCAN_STYLES.confirmIcon}>
                    <Icon iconName="Accept" />
                  </div>
                  <Text variant="mediumPlus" block>
                    Movement Recorded
                  </Text>
                  <Text variant="small" style={{ color: theme.palette.neutralSecondary, marginTop: 4 }}>
                    {movementDetail}
                  </Text>
                  <Text variant="xSmall" style={{ color: theme.palette.neutralTertiaryAlt, marginTop: 6 }}>
                    Lists updated
                  </Text>
                  <DefaultButton
                    text="Scan Next"
                    iconProps={{ iconName: 'QRCode' }}
                    onClick={resetScan}
                    disabled={saving}
                    styles={{ root: { marginTop: 14 } }}
                  />
                </div>
              ) : selectedMaterial ? (
                <div className={QR_SCAN_STYLES.resultCard}>
                  <div className={QR_SCAN_STYLES.resultTop}>
                    <Icon iconName="Accept" /> Material Identified
                  </div>
                  <div className={QR_SCAN_STYLES.resultBody}>
                    <div className={QR_SCAN_STYLES.materialCode}>{selectedMaterial.Material_Code}</div>
                    <div className={QR_SCAN_STYLES.materialDesc}>
                      {[selectedMaterial.Size, selectedMaterial.Material_Name, selectedMaterial.Specification]
                        .filter(Boolean)
                        .join(' • ')}
                    </div>

                    <div className={QR_SCAN_STYLES.detailGrid}>
                      <div className={QR_SCAN_STYLES.detailItem}>
                        <div className={QR_SCAN_STYLES.detailLabel}>Category</div>
                        <div className={QR_SCAN_STYLES.detailValue}>
                          <span className={`${QR_SCAN_STYLES.tag} ${QR_SCAN_STYLES[getCategoryTagClass(selectedMaterial.Category)]}`}>
                            {selectedMaterial.Category || 'N/A'}
                          </span>
                        </div>
                      </div>
                      <div className={QR_SCAN_STYLES.detailItem}>
                        <div className={QR_SCAN_STYLES.detailLabel}>Size</div>
                        <div className={QR_SCAN_STYLES.detailValue}>{selectedMaterial.Size || 'N/A'}</div>
                      </div>
                      <div className={QR_SCAN_STYLES.detailItem}>
                        <div className={QR_SCAN_STYLES.detailLabel}>UOM</div>
                        <div className={QR_SCAN_STYLES.detailValue}>{selectedMaterial.UOM || 'N/A'}</div>
                      </div>
                      <div className={QR_SCAN_STYLES.detailItem}>
                        <div className={QR_SCAN_STYLES.detailLabel}>Warehouse</div>
                        <div className={QR_SCAN_STYLES.detailValue}>{toLocation || fromLocation || 'Select location'}</div>
                      </div>
                      <div className={QR_SCAN_STYLES.detailItem}>
                        <div className={QR_SCAN_STYLES.detailLabel}>Movement</div>
                        <div className={QR_SCAN_STYLES.detailValue}>
                          <span className={`${QR_SCAN_STYLES.tag} ${QR_SCAN_STYLES[getMovementTagClass(movementType)]}`}>
                            {movementType}
                          </span>
                        </div>
                      </div>
                      <div className={QR_SCAN_STYLES.detailItem}>
                        <div className={QR_SCAN_STYLES.detailLabel}>Project</div>
                        <div className={QR_SCAN_STYLES.detailValue}>{projectCode || 'Not assigned'}</div>
                      </div>
                    </div>

                    {qrCodeUrl && (
                      <div className={QR_SCAN_STYLES.qrPanel}>
                        <div className={QR_SCAN_STYLES.qrBox}>
                          <Image
                            src={qrCodeUrl}
                            imageFit={ImageFit.contain}
                            width={52}
                            height={52}
                            alt="Material QR code"
                          />
                        </div>
                        <div className={QR_SCAN_STYLES.qrInfo}>
                          <div className={QR_SCAN_STYLES.qrCode}>{selectedMaterial.Material_Code}</div>
                          <div className={QR_SCAN_STYLES.qrDesc}>{selectedMaterial.Material_Name}</div>
                          <div className={QR_SCAN_STYLES.qrLoc}>{getMaterialQRCodeUrl(selectedMaterial) ? 'Stored QR Code URL' : 'Generated by QRCodeService'}</div>
                        </div>
                      </div>
                    )}

                    <div className={QR_SCAN_STYLES.actionTitle}>Select Action</div>
                    <div className={QR_SCAN_STYLES.actionGrid}>
                      {MOVEMENT_TYPES.map((movement) => (
                        <div
                          key={movement.key}
                          className={`${QR_SCAN_STYLES.actionCard} ${
                            movementType === movement.key ? QR_SCAN_STYLES.actionCardSelected : QR_SCAN_STYLES.actionCardHover
                          }`}
                          onClick={() => {
                            setMovementType(movement.key);
                            setConfirmedMovement(null);
                          }}
                        >
                          <Icon iconName={movement.icon} className={QR_SCAN_STYLES.actionCardIcon} />
                          <span className={QR_SCAN_STYLES.actionCardText}>{movement.label}</span>
                        </div>
                      ))}
                    </div>

                    <div className={QR_SCAN_STYLES.actionControls}>
                      <Dropdown
                        label="Location"
                        placeholder="Select warehouse"
                        selectedKey={fromLocation || toLocation || undefined}
                        options={warehouseOptionsForDropdown}
                        onChange={(_, option) => {
                          const nextLocation = (option?.key as string) || '';
                          if (movementType === 'GRN' || movementType === 'Return') {
                            setToLocation(nextLocation);
                          } else {
                            setFromLocation(nextLocation);
                          }
                        }}
                        disabled={saving}
                        className={QR_SCAN_STYLES.compactDropdown}
                      />
                      <TextField
                        label="Qty"
                        type="number"
                        min="0.0001"
                        step="0.0001"
                        value={quantity}
                        onChange={(_, value) => setQuantity(value || '')}
                        disabled={saving}
                        className={QR_SCAN_STYLES.compactTextField}
                      />
                      <TextField
                        label="Project code"
                        value={projectCode}
                        onChange={(_, value) => setProjectCode(value || '')}
                        disabled={saving}
                        styles={{ root: { minWidth: 150 } }}
                      />
                      <PrimaryButton
                        text="Confirm"
                        iconProps={{ iconName: 'Accept' }}
                        onClick={() => handleSaveMovement().catch(() => undefined)}
                        disabled={!canSave}
                      />
                    </div>

                    <TextField
                      label="Verification note"
                      multiline
                      autoAdjustHeight
                      value={note}
                      onChange={(_, value) => setNote(value || '')}
                      disabled={saving}
                      styles={{ root: { marginTop: 12 } }}
                    />
                  </div>
                </div>
              ) : (
                <div className={QR_SCAN_STYLES.emptyCard}>
                  <Icon
                    iconName="QRCode"
                    style={{ fontSize: 48, color: theme.palette.neutralSecondary, marginBottom: 12 }}
                  />
                  <Text variant="mediumPlus" block>
                    Scan a material QR code
                  </Text>
                  <Text variant="small" style={{ color: theme.palette.neutralSecondary, marginTop: 4 }}>
                    Use the camera or paste the generated material code to view item details.
                  </Text>
                </div>
              )}
            </div>
          </div>
        )}
      </Stack>
    </>
  );
};

export default QRScanMovePanel;
export { parseMaterialCodeFromScan };
