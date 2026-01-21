export interface TextItem {
  text: string;
  x: number; // inches
  y: number; // inches
  w: number; // inches
  h: number; // inches
  fontSize: number; // points
  fontFace: string;
  color?: string; // hex
  rotation?: number; // degrees
}

export interface ImageItem {
  data: string; // Base64
  x: number; // inches
  y: number; // inches
  w: number; // inches
  h: number; // inches
}

export interface ProcessedPage {
  pageNumber: number;
  imageData: string; // Base64 (Background only, vectors/shapes)
  width: number; // inches
  height: number; // inches
  textItems: TextItem[];
  images: ImageItem[];
  aiNotes?: string;
}

export enum ProcessingStatus {
  IDLE = 'IDLE',
  READING_PDF = 'READING_PDF',
  ANALYZING_AI = 'ANALYZING_AI',
  GENERATING_PPTX = 'GENERATING_PPTX',
  COMPLETED = 'COMPLETED',
  ERROR = 'ERROR'
}

export interface AppState {
  file: File | null;
  status: ProcessingStatus;
  progress: number; // 0-100
  processedPages: ProcessedPage[];
  errorMessage?: string;
  useAI: boolean;
  apiKey?: string;
  showApiKeyModal?: boolean;
}