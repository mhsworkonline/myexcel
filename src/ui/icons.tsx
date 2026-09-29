// Original Fluent-style SVG icons for large ribbon buttons (not Microsoft assets).
const G = '#107C41';
const B = '#2B7CD3';

export const PasteIcon = () => (
  <svg width="30" height="32" viewBox="0 0 30 32">
    <rect x="3" y="4" width="18" height="24" rx="2" fill="#C8A165" />
    <rect x="8" y="2" width="8" height="5" rx="1.5" fill="#6B6B6B" />
    <rect x="11" y="11" width="16" height="19" rx="1.5" fill="#fff" stroke="#6B6B6B" />
    <path d="M14 16h10M14 19.5h10M14 23h7" stroke="#8A8A8A" strokeWidth="1.2" />
  </svg>
);

export const CondFormatIcon = () => (
  <svg width="30" height="30" viewBox="0 0 30 30">
    <rect x="3" y="4" width="22" height="22" fill="#fff" stroke="#8A8A8A" />
    <path d="M3 11.3h22M3 18.6h22M10.3 4v22M17.6 4v22" stroke="#C8C8C8" />
    <rect x="4" y="5" width="6" height="6" fill="#F4B183" />
    <rect x="11" y="12" width="6" height="6" fill="#E53935" />
    <rect x="18" y="19" width="6" height="6" fill="#70AD47" />
    <rect x="18" y="5" width="6" height="6" fill="#FFD966" />
  </svg>
);

export const FormatTableIcon = () => (
  <svg width="30" height="30" viewBox="0 0 30 30">
    <rect x="3" y="4" width="24" height="22" fill="#fff" stroke="#8A8A8A" />
    <rect x="3" y="4" width="24" height="5" fill={B} />
    <rect x="3" y="13.5" width="24" height="4" fill="#BDD7EE" />
    <rect x="3" y="21.5" width="24" height="4.5" fill="#BDD7EE" />
    <path d="M11 9v17M19 9v17" stroke="#9DC3E6" />
  </svg>
);

export const CellStylesIcon = () => (
  <svg width="30" height="30" viewBox="0 0 30 30">
    <rect x="3" y="5" width="11" height="8" fill="#C6EFCE" stroke="#8A8A8A" />
    <rect x="16" y="5" width="11" height="8" fill="#FFC7CE" stroke="#8A8A8A" />
    <rect x="3" y="16" width="11" height="8" fill="#FFEB9C" stroke="#8A8A8A" />
    <rect x="16" y="16" width="11" height="8" fill={B} stroke="#8A8A8A" />
  </svg>
);

export const InsertCellsIcon = () => (
  <svg width="30" height="30" viewBox="0 0 30 30">
    <rect x="4" y="6" width="22" height="18" fill="#fff" stroke="#8A8A8A" />
    <path d="M4 12h22M4 18h22M11 6v18M18 6v18" stroke="#C8C8C8" />
    <rect x="11" y="12" width="7" height="6" fill="#BDD7EE" stroke={B} />
    <path d="M24 2v8M20 6h8" stroke={G} strokeWidth="2.2" />
  </svg>
);

export const DeleteCellsIcon = () => (
  <svg width="30" height="30" viewBox="0 0 30 30">
    <rect x="4" y="6" width="22" height="18" fill="#fff" stroke="#8A8A8A" />
    <path d="M4 12h22M4 18h22M11 6v18M18 6v18" stroke="#C8C8C8" />
    <rect x="11" y="12" width="7" height="6" fill="#F8CBAD" stroke="#C00000" />
    <path d="M21 3l6 6M27 3l-6 6" stroke="#C00000" strokeWidth="2" />
  </svg>
);

export const FormatCellsIcon = () => (
  <svg width="30" height="30" viewBox="0 0 30 30">
    <rect x="4" y="6" width="22" height="18" fill="#fff" stroke="#8A8A8A" />
    <path d="M4 12h22M4 18h22M11 6v18M18 6v18" stroke="#C8C8C8" />
    <rect x="11" y="12" width="7" height="6" fill="#E2EFDA" stroke={G} />
    <path d="M9 26l8-8 3 3-8 8H9z" fill="#FFC000" stroke="#7F6000" strokeWidth=".8" />
  </svg>
);

export const SortFilterIcon = () => (
  <svg width="30" height="30" viewBox="0 0 30 30">
    <text x="3" y="12" fontSize="9" fontWeight="700" fill={B} fontFamily="Segoe UI, Arial">A</text>
    <text x="3" y="24" fontSize="9" fontWeight="700" fill="#555" fontFamily="Segoe UI, Arial">Z</text>
    <path d="M13 5v18M10 20l3 4 3-4" stroke="#555" strokeWidth="1.6" fill="none" />
    <path d="M17 7h11l-4.2 5.6V20l-2.6 1.6v-9z" fill="#fff" stroke="#555" strokeWidth="1.2" />
  </svg>
);

export const FindSelectIcon = () => (
  <svg width="30" height="30" viewBox="0 0 30 30">
    <circle cx="12" cy="12" r="7" fill="#fff" stroke="#555" strokeWidth="2" />
    <path d="M17 17l8 8" stroke="#555" strokeWidth="3" strokeLinecap="round" />
  </svg>
);

export const PivotIcon = () => (
  <svg width="30" height="30" viewBox="0 0 30 30">
    <rect x="3" y="4" width="24" height="22" fill="#fff" stroke="#8A8A8A" />
    <rect x="3" y="4" width="24" height="5" fill="#E2EFDA" />
    <rect x="3" y="4" width="7" height="22" fill="#E2EFDA" />
    <path d="M10 4v22M3 9h24" stroke={G} />
    <path d="M15 14h8l-2-2M23 14l-2 2M19 18v6l-2-2M19 24l2-2" stroke={G} strokeWidth="1.3" fill="none" />
  </svg>
);

export const TableIcon = () => (
  <svg width="30" height="30" viewBox="0 0 30 30">
    <rect x="3" y="4" width="24" height="22" fill="#fff" stroke="#8A8A8A" />
    <rect x="3" y="4" width="24" height="5" fill={B} />
    <path d="M3 13.5h24M3 18h24M3 22h24M11 9v17M19 9v17" stroke="#9DC3E6" />
  </svg>
);

export const ChartColumnIcon = () => (
  <svg width="30" height="30" viewBox="0 0 30 30">
    <path d="M4 26h23" stroke="#8A8A8A" />
    <rect x="6" y="14" width="4" height="12" fill={B} />
    <rect x="12" y="8" width="4" height="18" fill="#ED7D31" />
    <rect x="18" y="17" width="4" height="9" fill="#A5A5A5" />
    <rect x="23" y="11" width="3" height="15" fill="#FFC000" />
  </svg>
);

export const RecommendedChartsIcon = () => (
  <svg width="30" height="30" viewBox="0 0 30 30">
    <rect x="3" y="5" width="20" height="16" fill="#fff" stroke="#8A8A8A" />
    <rect x="6" y="13" width="3" height="7" fill={B} />
    <rect x="11" y="9" width="3" height="11" fill="#ED7D31" />
    <rect x="16" y="15" width="3" height="5" fill="#70AD47" />
    <path d="M24 16l1.6 3.3 3.6.5-2.6 2.5.6 3.6-3.2-1.7-3.2 1.7.6-3.6-2.6-2.5 3.6-.5z" fill="#FFC000" stroke="#BF9000" strokeWidth=".6" />
  </svg>
);

export const AutoSumIcon = ({ size = 16 }: { size?: number }) => (
  <svg width={size} height={size} viewBox="0 0 16 16">
    <path d="M12.5 3V2h-9l5 6-5 6h9v-1" fill="none" stroke="#444" strokeWidth="1.4" strokeLinejoin="round" />
  </svg>
);

export const MergeIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16">
    <rect x="1" y="3" width="14" height="10" fill="none" stroke="#666" />
    <path d="M3 8h10M3 8l2-2M3 8l2 2M13 8l-2-2M13 8l-2 2" stroke={B} strokeWidth="1.2" fill="none" />
  </svg>
);

export const BordersIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16">
    <rect x="2" y="2" width="12" height="12" fill="none" stroke="#999" strokeDasharray="1 1" />
    <path d="M8 2v12M2 8h12" stroke="#999" strokeDasharray="1 1" />
    <path d="M2 14h12" stroke="#222" strokeWidth="1.6" />
  </svg>
);

export const FillBucketIcon = () => (
  <svg width="16" height="14" viewBox="0 0 16 14">
    <path d="M3 6.5L8 1.5l5 5-5 5z" fill="#fff" stroke="#444" />
    <path d="M13 8c0 0 1.8 2.2 1.8 3.2a1.8 1.8 0 01-3.6 0C11.2 10.2 13 8 13 8z" fill="#444" />
  </svg>
);

export const FontColorIcon = () => (
  <svg width="16" height="14" viewBox="0 0 16 14">
    <text x="3" y="11" fontSize="12" fontWeight="600" fontFamily="Segoe UI, Arial" fill="#333">A</text>
  </svg>
);

export const AppIcon = () => (
  <svg width="20" height="20" viewBox="0 0 32 32">
    <rect x="2" y="2" width="28" height="28" rx="6" fill={G} />
    <path d="M8 9h16M8 16h16M8 23h16M13 6v20M20 6v20" stroke="#fff" strokeWidth="2" opacity=".85" />
    <rect x="13" y="16" width="7" height="7" fill="#fff" />
  </svg>
);
