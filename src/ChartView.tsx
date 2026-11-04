import React from 'react';
import { SheetChart, Sheet, parseRange } from './model';
import { Trash2, Settings2 } from 'lucide-react';

interface ChartViewProps {
  chart: SheetChart;
  sheet: Sheet;
  onUpdate: (updated: SheetChart) => void;
  onDelete: (id: string) => void;
  onEdit: (chart: SheetChart) => void;
}

const PALETTE = ['#29967b', '#3b82f6', '#f59e0b', '#8b5cf6', '#ef4444', '#10b981', '#6366f1'];

export function ChartView({ chart, sheet, onUpdate, onDelete, onEdit }: ChartViewProps) {
  // Extract data from sheet range
  const range = parseRange(chart.range);
  const cellData = sheet.cellData || {};

  const labels: string[] = [];
  const seriesData: { name: string; values: number[]; color: string }[] = [];

  // Determine series names
  const firstRow = cellData[range.startRow] || {};
  const hasLabelsInCol = range.startCol < range.endCol;

  const dataStartRow = range.startRow + 1;
  const dataStartCol = hasLabelsInCol ? range.startCol + 1 : range.startCol;

  // Read series names from header row
  for (let c = dataStartCol; c <= range.endCol; c++) {
    const sName = String(firstRow[c]?.v ?? `Series ${c - dataStartCol + 1}`);
    const color = chart.series?.[c - dataStartCol]?.color || PALETTE[(c - dataStartCol) % PALETTE.length];
    seriesData.push({ name: sName, values: [], color });
  }

  // Read row data
  for (let r = dataStartRow; r <= range.endRow; r++) {
    const row = cellData[r] || {};
    const label = hasLabelsInCol ? String(row[range.startCol]?.v ?? `Row ${r + 1}`) : `Item ${r - dataStartRow + 1}`;
    labels.push(label);

    seriesData.forEach((s, sIdx) => {
      const colIdx = dataStartCol + sIdx;
      const rawVal = row[colIdx]?.v;
      const numVal = typeof rawVal === 'number' ? rawVal : parseFloat(String(rawVal ?? 0));
      s.values.push(Number.isFinite(numVal) ? numVal : 0);
    });
  }

  // SVG Geometry
  const width = chart.width || 480;
  const height = chart.height || 280;
  const padding = { top: 25, right: 30, bottom: 85, left: 55 };
  const plotWidth = Math.max(50, width - padding.left - padding.right);
  const plotHeight = Math.max(50, height - padding.top - padding.bottom);

  // Compute scale
  let allValues = seriesData.flatMap(s => s.values);
  if (allValues.length === 0) allValues = [0, 10];
  const minVal = Math.min(0, ...allValues);
  const maxVal = Math.max(1, ...allValues);
  const rangeSpan = maxVal - minVal || 1;

  const getY = (val: number) => padding.top + plotHeight - ((val - minVal) / rangeSpan) * plotHeight;
  const zeroY = getY(0);

  return (
    <div
      className="embedded-sheet-chart"
      style={{
        width: `${width}px`,
        height: `${height}px`,
        position: 'relative',
        background: 'var(--chart-bg, #ffffff)',
        border: '1px solid var(--border-color, #e2e8f0)',
        borderRadius: '8px',
        boxShadow: '0 4px 12px rgba(0,0,0,0.08)',
        overflow: 'hidden',
        margin: '10px 0',
      }}
    >
      <div
        className="chart-header-bar"
        style={{
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          padding: '6px 12px',
          background: 'var(--chart-header-bg, #f8fafc)',
          borderBottom: '1px solid var(--border-color, #e2e8f0)',
          fontSize: '13px',
          fontWeight: 600,
        }}
      >
        <span>{chart.title || 'Chart'}</span>
        <div style={{ display: 'flex', gap: '6px' }}>
          <button
            className="icon-btn"
            style={{ border: 'none', background: 'transparent', cursor: 'pointer', padding: '2px' }}
            title="Chart options"
            onClick={() => onEdit(chart)}
          >
            <Settings2 size={15} />
          </button>
          <button
            className="icon-btn"
            style={{ border: 'none', background: 'transparent', cursor: 'pointer', padding: '2px', color: '#ef4444' }}
            title="Delete chart"
            onClick={() => onDelete(chart.id)}
          >
            <Trash2 size={15} />
          </button>
        </div>
      </div>

      <svg width={width} height={height - 35} viewBox={`0 0 ${width} ${height - 35}`}>
        {/* Gridlines */}
        {chart.chartType !== 'pie' && chart.showGrid !== false && (
          <g className="gridlines" stroke="#e2e8f0" strokeDasharray="3 3">
            {[0, 0.25, 0.5, 0.75, 1].map((pct, i) => {
              const y = padding.top + plotHeight * (1 - pct);
              const val = minVal + rangeSpan * pct;
              return (
                <g key={i}>
                  <line x1={padding.left} y1={y} x2={padding.left + plotWidth} y2={y} />
                  <text
                    x={padding.left - 8}
                    y={y + 3}
                    textAnchor="end"
                    fontSize="10"
                    fill="#64748b"
                  >
                    {Math.round(val * 100) / 100}
                  </text>
                </g>
              );
            })}
          </g>
        )}

        {/* X Axis Baseline */}
        {chart.chartType !== 'pie' && <line
          x1={padding.left}
          y1={zeroY}
          x2={padding.left + plotWidth}
          y2={zeroY}
          stroke="#94a3b8"
          strokeWidth="1.5"
        />}

        {/* Render Chart Types */}
        {chart.chartType === 'column' && (
          <g className="chart-columns">
            {labels.map((label, rIdx) => {
              const bandWidth = plotWidth / labels.length;
              const barWidth = Math.max(4, (bandWidth - 8) / seriesData.length);
              const groupX = padding.left + rIdx * bandWidth + 4;

              return (
                <g key={rIdx}>
                  {seriesData.map((s, sIdx) => {
                    const val = s.values[rIdx] ?? 0;
                    const barY = val >= 0 ? getY(val) : zeroY;
                    const barHeight = Math.abs(getY(val) - zeroY);
                    const barX = groupX + sIdx * barWidth;

                    return (
                      <rect
                        key={sIdx}
                        x={barX}
                        y={barY}
                        width={Math.max(2, barWidth - 2)}
                        height={Math.max(1, barHeight)}
                        fill={s.color}
                        rx="2"
                      >
                        <title>{`${s.name} (${label}): ${val}`}</title>
                      </rect>
                    );
                  })}
                  <text
                    x={groupX + bandWidth / 2 - 4}
                    y={padding.top + plotHeight + 18}
                    textAnchor="middle"
                    fontSize="10"
                    fill="#64748b"
                  >
                    {label.length > 8 ? `${label.slice(0, 8)}…` : label}
                  </text>
                </g>
              );
            })}
          </g>
        )}

        {chart.chartType === 'bar' && (
          <g className="chart-bars">
            {labels.map((label, rIdx) => {
              const bandHeight = plotHeight / labels.length;
              const barHeight = Math.max(4, (bandHeight - 6) / seriesData.length);
              const groupY = padding.top + rIdx * bandHeight + 3;

              return (
                <g key={rIdx}>
                  {seriesData.map((s, sIdx) => {
                    const val = s.values[rIdx] ?? 0;
                    const barWidth = Math.abs(val) / rangeSpan * plotWidth;
                    const barY = groupY + sIdx * barHeight;

                    return (
                      <rect
                        key={sIdx}
                        x={padding.left + (Math.min(0, val) - minVal) / rangeSpan * plotWidth}
                        y={barY}
                        width={Math.max(1, barWidth)}
                        height={Math.max(2, barHeight - 2)}
                        fill={s.color}
                        rx="2"
                      >
                        <title>{`${s.name} (${label}): ${val}`}</title>
                      </rect>
                    );
                  })}
                  <text
                    x={padding.left - 8}
                    y={groupY + bandHeight / 2}
                    textAnchor="end"
                    fontSize="10"
                    fill="#64748b"
                  >
                    {label.length > 8 ? `${label.slice(0, 8)}…` : label}
                  </text>
                </g>
              );
            })}
          </g>
        )}

        {(chart.chartType === 'line' || chart.chartType === 'area') && (
          <g className="chart-lines">
            {seriesData.map((s, sIdx) => {
              const points = s.values.map((val, rIdx) => {
                const step = labels.length > 1 ? plotWidth / (labels.length - 1) : plotWidth / 2;
                const x = padding.left + rIdx * step;
                const y = getY(val);
                return { x, y, val, label: labels[rIdx] };
              });

              const dLine = points.map((p, i) => `${i === 0 ? 'M' : 'L'} ${p.x} ${p.y}`).join(' ');
              const dArea =
                dLine +
                ` L ${points[points.length - 1]?.x ?? padding.left} ${zeroY} L ${points[0]?.x ?? padding.left} ${zeroY} Z`;

              return (
                <g key={sIdx}>
                  {chart.chartType === 'area' && (
                    <path d={dArea} fill={s.color} fillOpacity="0.25" stroke="none" />
                  )}
                  <path d={dLine} fill="none" stroke={s.color} strokeWidth="2.5" />
                  {points.map((p, pIdx) => (
                    <circle key={pIdx} cx={p.x} cy={p.y} r="3.5" fill={s.color} stroke="#fff" strokeWidth="1.5">
                      <title>{`${s.name} (${p.label}): ${p.val}`}</title>
                    </circle>
                  ))}
                </g>
              );
            })}
            {labels.map((label, rIdx) => {
              const step = labels.length > 1 ? plotWidth / (labels.length - 1) : plotWidth / 2;
              const x = padding.left + rIdx * step;
              return (
                <text
                  key={rIdx}
                  x={x}
                  y={padding.top + plotHeight + 18}
                  textAnchor="middle"
                  fontSize="10"
                  fill="#64748b"
                >
                  {label.length > 8 ? `${label.slice(0, 8)}…` : label}
                </text>
              );
            })}
          </g>
        )}

        {chart.chartType === 'pie' && (
          <g className="chart-pie">
            {(() => {
              const s = seriesData[0];
              if (!s) return null;
              const total = s.values.reduce((sum, v) => sum + Math.max(0, v), 0) || 1;
              const cx = padding.left + plotWidth / 2;
              const cy = padding.top + plotHeight / 2;
              const r = Math.min(plotWidth, plotHeight) / 2 - 10;
              let startAngle = 0;

              return s.values.map((val, idx) => {
                const sliceVal = Math.max(0, val);
                const angle = (sliceVal / total) * 2 * Math.PI;
                const endAngle = startAngle + angle;

                const x1 = cx + r * Math.sin(startAngle);
                const y1 = cy - r * Math.cos(startAngle);
                const x2 = cx + r * Math.sin(endAngle);
                const y2 = cy - r * Math.cos(endAngle);

                const largeArc = angle > Math.PI ? 1 : 0;
                const pathD = `M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${largeArc} 1 ${x2} ${y2} Z`;
                const sliceColor = PALETTE[idx % PALETTE.length];

                startAngle = endAngle;
                if (sliceVal >= total && sliceVal > 0) return <circle key={idx} cx={cx} cy={cy} r={r} fill={sliceColor}><title>{labels[idx]}: {val}</title></circle>;

                return (
                  <path key={idx} d={pathD} fill={sliceColor} stroke="#ffffff" strokeWidth="1.5">
                    <title>{`${labels[idx] || `Item ${idx + 1}`}: ${val} (${Math.round((val / total) * 100)}%)`}</title>
                  </path>
                );
              });
            })()}
          </g>
        )}

        {/* Legend */}
        {chart.showLegend !== false && seriesData.length > 0 && (
          <g
            className="chart-legend"
            transform={`translate(${padding.left}, ${height - 60})`}
          >
            {seriesData.map((s, idx) => (
              <g key={idx} transform={`translate(${idx * 110}, 0)`}>
                <rect x="0" y="0" width="12" height="12" rx="2" fill={s.color} />
                <text x="18" y="10" fontSize="11" fill="#475569">
                  {s.name.length > 12 ? `${s.name.slice(0, 12)}…` : s.name}
                </text>
              </g>
            ))}
          </g>
        )}
      </svg>
    </div>
  );
}
