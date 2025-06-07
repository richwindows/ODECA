import React, { useState } from 'react';
import * as XLSX from 'xlsx';

const PROCESSED_VIEW_HEADERS = [
    'Order No', 'Type', 'Material Code', 'Material Position', 'Color', '(W/H)', 
    'Length', 'Angles', 'Qty', 'Cart No', 'Material Length', 'Cutting ID', 
    'Pieces ID', 'Window Number', 'Grid', 'LOCK', 'Nailing Fin'
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
    // 寻找最佳组合的函数
    function findBestCombination(items, maxLength) {
        let bestCombination = { items: [], totalLength: 0, waste: maxLength };
        
        // 生成所有可能的组合
        function generateCombinations(itemList, currentCombination, currentLength, startIndex) {
            // 检查当前组合是否更优
            if (currentLength <= maxLength) {
                const waste = maxLength - currentLength;
                if (waste < bestCombination.waste || 
                    (waste === bestCombination.waste && currentCombination.length > bestCombination.items.length)) {
                    bestCombination = {
                        items: [...currentCombination],
                        totalLength: currentLength,
                        waste: waste
                    };
                }
            }
            
            // 继续尝试添加更多项目
            for (let i = startIndex; i < itemList.length; i++) {
                const item = itemList[i];
                const itemTotalLength = parseFloat(item['Length']) * item.mergedCount;
                
                if (currentLength + itemTotalLength <= maxLength) {
                    currentCombination.push(item);
                    generateCombinations(itemList, currentCombination, currentLength + itemTotalLength, i + 1);
                    currentCombination.pop();
                }
            }
        }
        
        generateCombinations(items, [], 0, 0);
        return bestCombination;
    }

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

        // --- New Logic: Pairs + Bin Packing on sorted data ---
        
        // 1. Group identical items and sum their quantities
        // 在数据分组时添加更严格的验证和日志
        const itemGroups = new Map();
        let skippedItems = [];
        
        sortedData.forEach(row => {
            // 验证关键字段
            const qty = parseInt(row['Qty'], 10);
            const length = parseFloat(row['Length']);
            
            if (!qty || qty <= 0) {
                skippedItems.push({...row, reason: 'Invalid Qty'});
                return;
            }
            
            if (!length || length <= 0) {
                skippedItems.push({...row, reason: 'Invalid Length'});
                return;
            }
            
            const key = [
                row['Type'], row['Material Code'], row['Material Position'], row['Color'],
                row['(W/H)'], row['Length'], row['Angles'], row['Window Number'],
                row['Grid'], row['LOCK'], row['Nailing Fin']
            ].join('|');
            
            if (!itemGroups.has(key)) {
                itemGroups.set(key, { ...row, Qty: 0 });
            }
            itemGroups.get(key).Qty += qty;
        });
        
        // 转换为数组并添加调试信息
        const groupedDataArray = Array.from(itemGroups.values());
        
        // 在控制台以表格形式显示分组后的数据
        console.log('=== 分组后数据统计 ===');
        console.log('原始数据总行数:', sortedData.length);
        console.log('分组后数据行数:', groupedDataArray.length);
        console.log('总数量:', groupedDataArray.reduce((sum, item) => sum + item.Qty, 0));
        
        if (skippedItems.length > 0) {
            console.warn('跳过的数据项数量:', skippedItems.length);
            console.table(skippedItems);
        }
        
        console.log('\n=== 分组后数据详情 ===');
        console.table(groupedDataArray.map(item => ({
            '类型': item['Type'],
            '材料代码': item['Material Code'],
            '材料位置': item['Material Position'],
            '颜色': item['Color'],
            '宽高': item['(W/H)'],
            '长度': item['Length'],
            '角度': item['Angles'],
            '数量': item['Qty'],
            '窗号': item['Window Number']
        })));
        
        // 按颜色和材料代码统计
        const materialStats = new Map();
        groupedDataArray.forEach(item => {
            const key = `${item['Color']}-${item['Material Code']}`;
            if (!materialStats.has(key)) {
                materialStats.set(key, { count: 0, totalQty: 0 });
            }
            materialStats.get(key).count += 1;
            materialStats.get(key).totalQty += item.Qty;
        });
        
        console.log('\n=== 按颜色和材料代码统计 ===');
        console.table(Array.from(materialStats.entries()).map(([key, stats]) => {
            const [color, materialCode] = key.split('-');
            return {
                '颜色': color,
                '材料代码': materialCode,
                '项目数': stats.count,
                '总数量': stats.totalQty
            };
        }));
        // 添加调试信息
        if (skippedItems.length > 0) {
            console.warn('跳过的数据项:', skippedItems);
        }

        console.log('原始数据总数:', sortedData.length);
        console.log('分组后数据总数:', itemGroups.size);
        console.log('总数量:', Array.from(itemGroups.values()).reduce((sum, item) => sum + item.Qty, 0));

        const pairedItems = [];
        const singleItemsToPack = [];
        let cartNoCounter = 1;
        let pairedPiecesIdCounter = 1;

        // 2. Create paired items (Qty=2) and collect single leftovers from each group
        // 2. 处理配对项目 (Qty=2) - 使用与qty=1相同的优化切割算法
        const pairedItemsByMaterial = new Map();
        itemGroups.forEach(group => {
            const totalQty = group.Qty;
            const numPairs = Math.floor(totalQty / 2);
            const remainder = totalQty % 2;
            
            const { Qty, 'Cart No': _, ...baseData } = group;

            // 将配对项目按材料代码和颜色分组
            if (numPairs > 0) {
                const key = `${baseData['Material Code']}|${baseData['Color']}`;
                if (!pairedItemsByMaterial.has(key)) {
                    pairedItemsByMaterial.set(key, []);
                }
                
                // 为每对创建一个项目 - 修复mergedCount
                for (let i = 0; i < numPairs; i++) {
                    pairedItemsByMaterial.get(key).push({
                        ...baseData,
                        mergedCount: 2 // 修复：每个配对项目应该是2个，不是1个
                    });
                }
            }

            if (remainder > 0) {
                singleItemsToPack.push(baseData);
            }
        });
        
        // 显示按qty分割后的数据
        console.log('\n=== 按Qty分割后的数据统计 ===');
        
        // 显示配对项目（Qty=2）的分割结果
        const pairedItemsFlat = [];
        pairedItemsByMaterial.forEach((items, materialKey) => {
            const [materialCode, color] = materialKey.split('|');
            items.forEach((item, index) => {
                pairedItemsFlat.push({
                    '分组': `配对项目-${materialCode}-${color}`,
                    '序号': index + 1,
                    '类型': item['Type'],
                    '材料代码': item['Material Code'],
                    '颜色': item['Color'],
                    '长度': item['Length'],
                    '原始Qty': '2',
                    '合并数量': item.mergedCount,
                    '窗号': item['Window Number']
                });
            });
        });
        
        console.log('配对项目总数:', pairedItemsFlat.length);
        if (pairedItemsFlat.length > 0) {
            console.table(pairedItemsFlat);
        }
        
        // 显示单个项目（Qty=1或余数）的分割结果
        const singleItemsDisplay = singleItemsToPack.map((item, index) => ({
            '分组': '单个项目',
            '序号': index + 1,
            '类型': item['Type'],
            '材料代码': item['Material Code'],
            '颜色': item['Color'],
            '长度': item['Length'],
            '原始Qty': '1',
            '窗号': item['Window Number']
        }));
        
        console.log('单个项目总数:', singleItemsDisplay.length);
        if (singleItemsDisplay.length > 0) {
            console.table(singleItemsDisplay);
        }
        
        // 按材料和颜色统计分割结果
        const splitStats = new Map();
        
        // 统计配对项目
        pairedItemsByMaterial.forEach((items, materialKey) => {
            const [materialCode, color] = materialKey.split('|');
            const key = `${color}-${materialCode}`;
            if (!splitStats.has(key)) {
                splitStats.set(key, { paired: 0, single: 0 });
            }
            splitStats.get(key).paired += items.length;
        });
        
        // 统计单个项目
        singleItemsToPack.forEach(item => {
            const key = `${item['Color']}-${item['Material Code']}`;
            if (!splitStats.has(key)) {
                splitStats.set(key, { paired: 0, single: 0 });
            }
            splitStats.get(key).single += 1;
        });
        
        console.log('\n=== 分割统计汇总 ===');
        console.table(Array.from(splitStats.entries()).map(([key, stats]) => {
            const [color, materialCode] = key.split('-');
            return {
                '颜色': color,
                '材料代码': materialCode,
                '配对项目数': stats.paired,
                '单个项目数': stats.single,
                '总项目数': stats.paired + stats.single
            };
        }));
        
        // 对配对项目应用完全相同的优化切割算法
        // 删除这行：const pairedItems = [];
        pairedItemsByMaterial.forEach(items => {
            // 简化合并逻辑，直接按长度分组而不进行复杂的二次合并
            const lengthGroups = new Map();
            
            items.forEach(item => {
                const lengthKey = item['Length'];
                if (!lengthGroups.has(lengthKey)) {
                    lengthGroups.set(lengthKey, {
                        ...item,
                        mergedCount: 0
                    });
                }
                // 直接累加mergedCount，确保数据不丢失
                lengthGroups.get(lengthKey).mergedCount += item.mergedCount;
            });
            
            const mergedItems = Array.from(lengthGroups.values());
            
            // 添加数据验证
            const originalTotal = items.reduce((sum, item) => sum + item.mergedCount, 0);
            const mergedTotal = mergedItems.reduce((sum, item) => sum + item.mergedCount, 0);
            
            console.log('\n=== 配对项目合并验证 ===');
            console.log('原始总数:', originalTotal);
            console.log('合并后总数:', mergedTotal);
            console.log('数据是否一致:', originalTotal === mergedTotal ? '✓' : '❌');
            
            if (originalTotal !== mergedTotal) {
                console.error('❌ 配对项目数据在合并过程中丢失！');
                console.table(items.map(item => ({
                    '材料代码': item['Material Code'],
                    '颜色': item['Color'],
                    '长度': item['Length'],
                    '合并数量': item.mergedCount
                })));
            }
            lengthGroups.forEach(item => mergedItems.push(item));
            
            // 按长度降序排序
            mergedItems.sort((a, b) => parseFloat(b['Length']) - parseFloat(a['Length']));

            
            
            const bars = [];
            const remainingItems = [...mergedItems];
            
            while (remainingItems.length > 0) {
                console.log('\n=== 处理剩余项目 ===');
                console.log('剩余项目数量:', remainingItems.length);
                
                const bestCombination = findBestCombination(remainingItems, 210);
                
                if (bestCombination.items.length === 0) {
                    // 如果没有找到合适的组合，取第一个项目单独处理
                    const item = remainingItems.shift();
                    const totalLength = parseFloat(item['Length']) * (item.mergedCount || 1);
                    
                    console.log('\n=== 单独处理项目 ===');
                    console.log('项目:', item);
                    console.log('总长度:', totalLength);
                    
                    if (totalLength <= 210) {
                        bars.push({
                            remainingLength: 210 - totalLength,
                            pieces: [{
                                ...item,
                                cuttingId: item.mergedCount || 1
                            }]
                        });
                        console.log('✓ 项目已添加到bars');
                    } else {
                        console.error('❌ 项目长度超过最大长度，无法处理:', item);
                        // 可以选择拆分或其他处理方式
                    }
                } else {
                    // 使用最佳组合创建新的材料条
                    console.log('\n=== 使用最佳组合 ===');
                    bars.push({
                        remainingLength: bestCombination.waste,
                        pieces: bestCombination.items.map(item => ({
                            ...item,
                            cuttingId: item.mergedCount || 1
                        }))
                    });
                    
                    // 从剩余项目中移除已使用的项目（添加详细日志）
                    const removedSingleItemsDisplay = [];
                    bestCombination.items.forEach(usedItem => {
                        const index = remainingItems.findIndex(item => 
                            item['Length'] === usedItem['Length'] && 
                            item['Material Code'] === usedItem['Material Code'] &&
                            item['Color'] === usedItem['Color'] &&
                            (item.mergedCount || 1) === (usedItem.mergedCount || 1)
                        );
                        if (index !== -1) {
                            removedSingleItemsDisplay.push({
                                '序号': removedSingleItemsDisplay.length + 1,
                                '类型': remainingItems[index]['Type'],
                                '材料代码': remainingItems[index]['Material Code'],
                                '颜色': remainingItems[index]['Color'],
                                '长度': remainingItems[index]['Length'],
                                '合并数量': remainingItems[index].mergedCount || 1,
                                '状态': '已移除'
                            });
                            remainingItems.splice(index, 1);
                        } else {
                            console.error('❌ 未找到要移除的项目:', usedItem);
                            console.error('当前剩余项目:', remainingItems);
                            removedSingleItemsDisplay.push({
                                '序号': removedSingleItemsDisplay.length + 1,
                                '类型': usedItem['Type'],
                                '材料代码': usedItem['Material Code'],
                                '颜色': usedItem['Color'],
                                '长度': usedItem['Length'],
                                '合并数量': usedItem.mergedCount || 1,
                                '状态': '未找到'
                            });
                        }
                    });
                    
                    // 显示移除记录
                    if (removedSingleItemsDisplay.length > 0) {
                        console.log('\n=== 单个项目移除记录 ===');
                        console.table(removedSingleItemsDisplay);
                    }
                }
                
                // 防止无限循环
                if (remainingItems.length > 0 && bestCombination.items.length === 0) {
                    console.error('❌ 检测到可能的无限循环，强制退出');
                    break;
                }
            }
            
            // 将结果添加到pairedItems
            bars.forEach((bar, barIndex) => {
                bar.pieces.forEach((piece, pieceIndex) => {
                    pairedItems.push({
                        ...piece,
                        'Qty': 2, // 保持原始的qty=2标识
                        'Cutting ID': piece.cuttingId,
                        'Pieces ID': pieceIndex + 1,
                        'Material Length': 210,
                        'Cart No': cartNoCounter++,
                        'LOCK': 1
                    });
                });
            });
        });
        
        // 3. Bin-pack the leftover single items
        // 3. 使用贪婪算法优化单个项目的切割
        const singleItemsByMaterial = new Map();
        singleItemsToPack.forEach(item => {
            const key = `${item['Material Code']}|${item['Color']}`; // 按材料代码和颜色分组
            if (!singleItemsByMaterial.has(key)) {
                singleItemsByMaterial.set(key, []);
            }
            singleItemsByMaterial.get(key).push(item);
        });
        
        const packedItems = [];
        singleItemsByMaterial.forEach(items => {
            // 先按长度分组，合并相同长度的项目
            const lengthGroups = new Map();
            items.forEach(item => {
                const length = parseFloat(item['Length']);
                if (!lengthGroups.has(length)) {
                    lengthGroups.set(length, []);
                }
                lengthGroups.get(length).push(item);
            });
            
            // 将相同长度的项目合并为一条记录
            const mergedItems = [];
            lengthGroups.forEach((itemsOfSameLength, length) => {
                if (itemsOfSameLength.length > 0) {
                    const firstItem = itemsOfSameLength[0];
                    mergedItems.push({
                        ...firstItem,
                        'Length': length,
                        mergedCount: itemsOfSameLength.length // 记录合并的数量
                    });
                }
            });
            
            // 按长度降序排序，优先处理长的材料
            mergedItems.sort((a, b) => parseFloat(b['Length']) - parseFloat(a['Length']));
            
            // 改进的切割优化算法
            const bars = [];
            const remainingItems = [...mergedItems];
            
            while (remainingItems.length > 0) {
                const bestCombination = findBestCombination(remainingItems, 210);
                
                if (bestCombination.items.length === 0) {
                    // 如果没有找到合适的组合，取第一个项目单独处理
                    const item = remainingItems.shift();
                    const totalLength = parseFloat(item['Length']) * item.mergedCount;
                    if (totalLength <= 210) {
                        bars.push({
                            remainingLength: 210 - totalLength,
                            pieces: [{
                                ...item,
                                cuttingId: item.mergedCount
                            }]
                        });
                    }
                } else {
                    // 使用最佳组合创建新的材料条
                    bars.push({
                        remainingLength: bestCombination.waste,
                        pieces: bestCombination.items.map(item => ({
                            ...item,
                            cuttingId: item.mergedCount
                        }))
                    });
                    
                    // 从剩余项目中移除已使用的项目
                    bestCombination.items.forEach(usedItem => {
                        const index = remainingItems.findIndex(item => 
                            item['Length'] === usedItem['Length'] && 
                            item['Material Code'] === usedItem['Material Code'] &&
                            item['Color'] === usedItem['Color']
                        );
                        if (index !== -1) {
                            remainingItems.splice(index, 1);
                        }
                    });
                }
            }
            
            // 将结果添加到packedItems，每根材料条的pieces id从1开始
            bars.forEach((bar, barIndex) => {
                bar.pieces.forEach((piece, pieceIndex) => {
                    packedItems.push({
                        ...piece,
                        'Qty': 1,
                        'Cutting ID': piece.cuttingId, // 记录相同长度的数量
                        'Pieces ID': pieceIndex + 1, // 每根材料条内从1开始排序
                        'Material Length': 210,
                        'Cart No': cartNoCounter++,
                        'LOCK': 1
                    });
                });
            });
        });

        // 4. Combine results and update state
        const finalProcessedData = [...pairedItems, ...packedItems];
        setProcessedData(finalProcessedData);
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