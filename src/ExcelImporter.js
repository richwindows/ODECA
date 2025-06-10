import React, { useState } from 'react';
import * as XLSX from 'xlsx';

const PROCESSED_VIEW_HEADERS = [
    'Order No', 'Type', 'Material Code', 'Material Position', 'Color', '(W/H)', 
    'Length', 'Angles', 'Qty', 'Cart No', 'Material Length', 'Cutting ID', 
    'Pieces ID', 'Window Number', 'LOCK'
];

const DISPLAY_HEADERS = [
    'Order No', 'Type', 'Material Code', 'Material Position', 'Color', '(W/H)', 
    'Length', 'Angles', 'Qty', 'Cart No', 'Material Length', 'Cutting ID', 
    'Pieces ID', 'Window Number'
];

function ExcelImporter() {
  const [processedData, setProcessedData] = useState([]);
  const [originalData, setOriginalData] = useState([]);
  const [originalHeaders, setOriginalHeaders] = useState([]);
  const [view, setView] = useState('processed'); // 'original' or 'processed'

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1 });
      
      if (jsonData.length > 0) {
        const fileHeaders = jsonData[0].map(h => typeof h === 'string' ? h.trim() : h);
        const fileData = jsonData.slice(1);

        const originalDataAsObjects = fileData.map(row => {
            const rowData = {};
            fileHeaders.forEach((header, index) => {
                rowData[header] = row[index];
            });
            return rowData;
        });
        setOriginalHeaders(fileHeaders);
        setOriginalData(originalDataAsObjects);

        // --- Sort data before processing ---
        const sortedData = [...originalDataAsObjects].sort((a, b) => {
          const colorA = a['Color'] || '';
          const colorB = b['Color'] || '';
          if (colorA.localeCompare(colorB) !== 0) return colorA.localeCompare(colorB);
    
          const materialCodeA = a['Material Code'] || '';
          const materialCodeB = b['Material Code'] || '';
          if (materialCodeA.localeCompare(materialCodeB) !== 0) return materialCodeA.localeCompare(materialCodeB);
    
          const lengthA = parseFloat(a['Length']) || 0;
          const lengthB = parseFloat(b['Length']) || 0;
          return lengthB - lengthA;
        });

        // 1. Group identical items and sum their quantities
        const itemGroups = new Map();
        const skippedItems = [];
        
        sortedData.forEach(item => {
            const qty = parseInt(item['Qty']) || 0;
            const length = parseFloat(item['Length']) || 0;
            
            // 验证数据有效性
            if (qty <= 0 || length <= 0) {
                skippedItems.push({
                    ...item,
                    reason: `无效数据: Qty=${qty}, Length=${length}`
                });
                return;
            }
            
            const key = `${item['Type']}|${item['Material Code']}|${item['Color']}|${item['Length']}|${item['Window Number']}`;
            
            if (itemGroups.has(key)) {
                itemGroups.get(key).Qty += qty;
            } else {
                itemGroups.set(key, { ...item, Qty: qty });
            }
        });
        
        if (skippedItems.length > 0) {
            console.warn('跳过的数据项:', skippedItems);
        }

        console.log('原始数据总数:', sortedData.length);
        console.log('分组后数据总数:', itemGroups.size);
        console.log('总数量:', Array.from(itemGroups.values()).reduce((sum, item) => sum + item.Qty, 0));

        // 2. 按qty拆分数据 - 拆分成数量为2和1的项目
        const processedItems = [];
        let cartNoCounter = 1;

        itemGroups.forEach(group => {
            const totalQty = group.Qty;
            const { Qty, 'Cart No': _, ...baseData } = group;

            // 拆分逻辑：优先创建数量为2的项目，剩余创建数量为1的项目
            let remainingQty = totalQty;
            
            // 创建数量为2的项目
            while (remainingQty >= 2) {
                processedItems.push({
                    ...baseData,
                    'Qty': 2,
                    'Material Length': 210, // 默认材料长度
                    'Cart No': cartNoCounter++,
                    'LOCK': 1
                });
                remainingQty -= 2;
            }
            
            // 如果还有剩余数量（必定是1），创建数量为1的项目
            if (remainingQty === 1) {
                processedItems.push({
                    ...baseData,
                    'Qty': 1,
                    'Material Length': 210, // 默认材料长度
                    'Cart No': cartNoCounter++,
                    'LOCK': 1
                });
            }
        });
        
        // 显示拆分结果统计
        console.log('\n=== 按Qty拆分结果统计 ===');
        console.log('拆分后项目总数:', processedItems.length);
        
        // 按材料、颜色和数量统计拆分结果
        const splitStats = new Map();
        const groupedByMaterialColorQty = new Map();
        
        processedItems.forEach(item => {
            // 原有的统计（按材料和颜色）
            const key = `${item['Color']}-${item['Material Code']}`;
            if (!splitStats.has(key)) {
                splitStats.set(key, 0);
            }
            splitStats.set(key, splitStats.get(key) + 1);
            
            // 新增：按材料、颜色和数量分组
            const groupKey = `${item['Color']}-${item['Material Code']}-Qty${item['Qty']}`;
            if (!groupedByMaterialColorQty.has(groupKey)) {
                groupedByMaterialColorQty.set(groupKey, []);
            }
            groupedByMaterialColorQty.get(groupKey).push(item);
        });
        
        console.log('\n=== 拆分统计汇总 ===');
        console.table(Array.from(splitStats.entries()).map(([key, count]) => {
            const [color, materialCode] = key.split('-');
            return {
                '颜色': color,
                '材料代码': materialCode,
                '拆分后数量': count
            };
        }));
        
        // 新增：按材料、颜色和数量分组的详细统计
        console.log('\n=== 按材料、颜色、数量分组统计 ===');
        const groupedStats = Array.from(groupedByMaterialColorQty.entries()).map(([key, items]) => {
            const [color, materialCode, qtyPart] = key.split('-');
            const qty = qtyPart.replace('Qty', '');
            return {
                '颜色': color,
                '材料代码': materialCode,
                '数量类型': `Qty=${qty}`,
                '项目数': items.length,
                '总数量': items.length * parseInt(qty)
            };
        }).sort((a, b) => {
            // 先按材料代码排序，再按颜色排序，最后按数量类型排序
            if (a['材料代码'] !== b['材料代码']) {
                return a['材料代码'].localeCompare(b['材料代码']);
            }
            if (a['颜色'] !== b['颜色']) {
                return a['颜色'].localeCompare(b['颜色']);
            }
            return a['数量类型'].localeCompare(b['数量类型']);
        });
        
        console.table(groupedStats);
        
        // 显示每个分组的详细信息（可选，用于调试）
        console.log('\n=== 分组详细信息 ===');
        groupedByMaterialColorQty.forEach((items, key) => {
            const [color, materialCode, qtyPart] = key.split('-');
            console.log(`\n${color} - ${materialCode} - ${qtyPart}: ${items.length}个项目`);
            if (items.length <= 5) {
                // 如果项目数量少于等于5个，显示所有项目
                console.table(items.map(item => ({
                    'Cart No': item['Cart No'],
                    '长度': item['Length'],
                    'Qty': item['Qty'],
                    '窗号': item['Window Number']
                })));
            } else {
                // 如果项目数量多于5个，只显示前3个和后2个
                console.log('显示前3个和后2个项目:');
                const sampleItems = [...items.slice(0, 3), ...items.slice(-2)];
                console.table(sampleItems.map(item => ({
                    'Cart No': item['Cart No'],
                    '长度': item['Length'],
                    'Qty': item['Qty'],
                    '窗号': item['Window Number']
                })));
            }
        });
        
        // 显示前10条拆分后的数据作为示例
        if (processedItems.length > 0) {
            console.log('\n=== 拆分后数据示例（前10条）===');
            const sampleData = processedItems.slice(0, 10).map((item, index) => ({
                '序号': index + 1,
                '类型': item['Type'],
                '材料代码': item['Material Code'],
                '颜色': item['Color'],
                '长度': item['Length'],
                'Qty': item['Qty'],
                'Cart No': item['Cart No'],
                'Pieces ID': item['Pieces ID'],
                '窗号': item['Window Number']
            }));
            console.table(sampleData);
        }

        // 3. 优化算法处理 - 不区分qty，全部统一处理（贪心法）
        console.log('\n=== 开始优化算法处理（所有项目统一处理）===');
        // 按材料代码和颜色分组
        const materialColorGroups = new Map();
        processedItems.forEach(item => {
            const groupKey = `${item['Material Code']}-${item['Color']}`;
            if (!materialColorGroups.has(groupKey)) {
                materialColorGroups.set(groupKey, []);
            }
            materialColorGroups.get(groupKey).push(item);
        });
        console.log(`材料颜色分组数: ${materialColorGroups.size}`);
        materialColorGroups.forEach((items, groupKey) => {
            let remainingItems = [...items].sort((a, b) => parseFloat(b.Length) - parseFloat(a.Length));
            let cuttingGroupCounter = 1;
            while (remainingItems.length > 0) {
                let currentGroup = [];
                let currentLength = 0;
                for (let i = 0; i < remainingItems.length; i++) {
                    const itemLength = parseFloat(remainingItems[i].Length);
                    if (currentLength + itemLength <= 210) {
                        currentGroup.push(remainingItems[i]);
                        currentLength += itemLength;
                    }
                }
                if (currentGroup.length === 0) {
                    // 单独处理
                    const singleItem = remainingItems.shift();
                    singleItem['Cutting Group'] = `${groupKey}-G${cuttingGroupCounter}`;
                    singleItem['Pieces ID'] = 1;
                    singleItem['Cutting ID'] = 1;
                    cuttingGroupCounter++;
                } else {
                    // 先按长度从小到大排序
                    const sortedGroup = [...currentGroup].sort((a, b) => parseFloat(a.Length) - parseFloat(b.Length));
                    // 统计同尺寸数量
                    const lengthGroups = {};
                    sortedGroup.forEach(item => {
                        const len = item.Length;
                        if (!lengthGroups[len]) lengthGroups[len] = [];
                        lengthGroups[len].push(item);
                    });
                    // 分配编号
                    let piecesId = 1;
                    Object.keys(lengthGroups).sort((a, b) => parseFloat(a) - parseFloat(b)).forEach(len => {
                        const group = lengthGroups[len];
                        group.forEach(item => {
                            item['Cutting Group'] = `${groupKey}-G${cuttingGroupCounter}`;
                            item['Pieces ID'] = piecesId++;
                            item['Cutting ID'] = group.length;
                        });
                    });
                    // 从剩余项目中移除已分组
                    currentGroup.forEach(usedItem => {
                        const index = remainingItems.findIndex(item =>
                            item['Cart No'] === usedItem['Cart No'] &&
                            item['Length'] === usedItem['Length'] &&
                            item['Window Number'] === usedItem['Window Number']
                        );
                        if (index !== -1) {
                            remainingItems.splice(index, 1);
                        }
                    });
                    cuttingGroupCounter++;
                }
            }
        });
        
        // 4. 对处理后的数据进行最终排序，确保同一切割组的项目连续显示
        console.log('\n=== 开始最终数据排序 ===');
        // 将所有处理后的数据按切割组和pieces id正序排序
        const finalSortedData = processedItems.sort((a, b) => {
            const groupA = a['Cutting Group'] || '';
            const groupB = b['Cutting Group'] || '';
            if (groupA !== groupB) {
                return groupA.localeCompare(groupB);
            }
            // 同组内按pieces id正序
            return (a['Pieces ID'] || 0) - (b['Pieces ID'] || 0);
        });

        // 合并同组同尺寸，只保留一行，pieces id重新分配，qty等字段不变
        const mergedData = [];
        let currentGroup = '';
        let groupRows = [];
        finalSortedData.forEach(item => {
            if (item['Cutting Group'] !== currentGroup) {
                // 新组，处理上一组
                if (groupRows.length > 0) {
                    const lengthMap = {};
                    groupRows.forEach(row => {
                        const len = row['Length'];
                        if (!lengthMap[len]) lengthMap[len] = row;
                    });
                    let pid = 1;
                    Object.values(lengthMap).forEach(row => {
                        row['Pieces ID'] = pid++;
                        mergedData.push(row);
                    });
                }
                groupRows = [];
                currentGroup = item['Cutting Group'];
            }
            groupRows.push(item);
        });
        // 处理最后一组
        if (groupRows.length > 0) {
            const lengthMap = {};
            groupRows.forEach(row => {
                const len = row['Length'];
                if (!lengthMap[len]) lengthMap[len] = row;
            });
            let pid = 1;
            Object.values(lengthMap).forEach(row => {
                row['Pieces ID'] = pid++;
                mergedData.push(row);
            });
        }

        // 用mergedData作为最终导出/展示数据
        // 重新分配Cart No，从1开始递增
        mergedData.forEach((row, idx) => {
            row['Cart No'] = idx + 1;
        });

        // Remove Cutting Group from final data as it is not needed in the output
        mergedData.forEach(row => delete row['Cutting Group']);
        
        setProcessedData(mergedData);
        setView('processed');
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleDownload = () => {
    const ws = XLSX.utils.json_to_sheet(processedData, { header: PROCESSED_VIEW_HEADERS });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "新数据表");
    XLSX.writeFile(wb, "processed-data.xlsx");
  };

  const renderTable = (tableHeaders, tableData) => (
    <table>
      <thead>
        <tr>
          {tableHeaders.map((header, index) => (
            <th key={index}>{header}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        {tableData.map((row, index) => (
          <tr key={index}>
            {tableHeaders.map((header, hIndex) => (
              <td key={hIndex}>{row[header]}</td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );

  return (
    <div>
      <input type="file" onChange={handleFileUpload} />
      {originalData.length > 0 && (
        <div className="view-switcher">
          <button onClick={() => setView('original')} disabled={view === 'original'}>
            原始数据
          </button>
          <button onClick={() => setView('processed')} disabled={view === 'processed'}>
            新数据表
          </button>
          {processedData.length > 0 && (
            <button onClick={handleDownload} style={{ marginLeft: '10px' }}>
              下载生成excel
            </button>
          )}
        </div>
      )}

      {view === 'original' && originalData.length > 0 && renderTable(originalHeaders, originalData)}
      {view === 'processed' && processedData.length > 0 && renderTable(DISPLAY_HEADERS, processedData)}
    </div>
  );
}

export default ExcelImporter;