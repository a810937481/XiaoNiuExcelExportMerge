// Excel余额汇总处理工具 - JavaScript文件
// 此文件包含所有数据处理逻辑和交互功能

// 当DOM内容加载完成后执行初始化
document.addEventListener('DOMContentLoaded', function() {
    // 获取DOM元素
    const fileInput1 = document.getElementById('file1');
    const fileInput2 = document.getElementById('file2');
    const fileName1 = document.getElementById('fileName1');
    const fileName2 = document.getElementById('fileName2');
    const processBtn = document.getElementById('processBtn');
    const resultSection = document.getElementById('resultSection');
    const downloadBtn = document.getElementById('downloadBtn');
    const progressArea = document.getElementById('progressArea');
    const progressBar = document.getElementById('progressBar');
    const preview1 = document.getElementById('preview1');
    const preview2 = document.getElementById('preview2');
    const resultPreview = document.getElementById('resultPreview');
    const consumptionStat = document.getElementById('consumptionStat');
    const depositStat = document.getElementById('depositStat');
    const balanceStat = document.getElementById('balanceStat');
    
    // 存储上传的文件和处理后的数据
    let file1 = null;
    let file2 = null;
    let processedData = null;
    let chart = null;
    let summaryData = [];
    let detailData = [];
    
    // 为文件上传区域设置事件监听
    fileInput1.addEventListener('change', (e) => handleFileUpload(e, fileName1, 1, preview1));
    fileInput2.addEventListener('change', (e) => handleFileUpload(e, fileName2, 2, preview2));
    
    /**
     * 处理文件上传事件
     * @param {Event} event - 文件上传事件
     * @param {HTMLElement} fileNameElement - 显示文件名的元素
     * @param {number} fileNum - 文件编号（1或2）
     * @param {HTMLElement} previewElement - 预览区域元素
     */
    function handleFileUpload(event, fileNameElement, fileNum, previewElement) {
        const file = event.target.files[0];
        if (file) {
            fileNameElement.textContent = file.name;
            fileNameElement.style.color = '#27ae60';
            
            if (fileNum === 1) file1 = file;
            if (fileNum === 2) file2 = file;
            
            checkFilesReady();
            
            // 读取文件并显示预览
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    // 使用SheetJS库读取Excel文件
                    const workbook = XLSX.read(data, {
                        type: 'array',
                        cellText: false,
                        cellDates: true,
                        dateNF: 'yyyy-mm-dd'
                    });
                    
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    
                    // 将工作表数据转换为JSON格式
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                        defval: "",
                        raw: false,
                        dateNF: 'yyyy-mm-dd'
                    });
                    
                    // 优化：将金额列转换为数字类型
                    jsonData.forEach(row => {
                        if (row['金额（元）'] !== undefined) {
                            // 尝试转换为数字，如果转换失败则置为0
                            const amount = parseFloat(row['金额（元）']);
                            row['金额（元）'] = isNaN(amount) ? 0 : amount;
                        }
                    });
                    
                    // 保存数据
                    if (fileNum === 1) summaryData = jsonData;
                    if (fileNum === 2) detailData = jsonData;
                    
                    // 显示数据预览
                    showDataPreview(jsonData, previewElement);
                    
                    // 计算并显示统计数据
                    calculateAndDisplayStats();
                } catch (error) {
                    console.error('文件读取错误:', error);
                    previewElement.innerHTML = `
                        <p style="color: #e74c3c;">
                            文件读取失败：${error.message || '未知错误'}。
                            请确保上传的是有效的Excel文件（.xlsx或.xls格式）
                        </p>
                    `;
                }
            };
            
            // 添加错误处理
            reader.onerror = function() {
                fileNameElement.textContent = '文件读取错误';
                fileNameElement.style.color = '#e74c3c';
                previewElement.innerHTML = '<p style="color: #e74c3c;">文件读取过程中发生错误，请重试</p>';
            };
            
            reader.readAsArrayBuffer(file);
        }
    }
    
    /**
     * 检查是否文件已上传完成
     * 优化：支持只上传一个文件的情况
     */
    function checkFilesReady() {
        // 只要file1存在就可以处理（file2可选）
        if (file1) {
            processBtn.disabled = false;
            progressArea.innerHTML = file2 
                ? '两个文件已就绪，可以开始处理数据！' 
                : '汇总表已就绪，可以开始处理数据！';
        } else {
            processBtn.disabled = true;
        }
    }
    
    /**
     * 显示数据预览
     * @param {Array} data - 要显示的数据
     * @param {HTMLElement} previewElement - 用于显示预览的元素
     */
    function showDataPreview(data, previewElement) {
        if (data.length === 0) {
            previewElement.innerHTML = '<p>文件中没有数据</p>';
            return;
        }
        
        // 只显示前5行数据
        const previewData = data.slice(0, 5);
        const headers = Object.keys(previewData[0]); // 获取表头
        
        let tableHtml = `
            <p>数据预览（前${previewData.length}行）：</p>
            <table>
                <thead>
                    <tr>
        `;
        
        // 添加表头
        headers.forEach(header => {
            tableHtml += `<th>${header}</th>`;
        });
        
        tableHtml += `</tr></thead><tbody>`;
        
        // 添加数据行
        previewData.forEach(row => {
            tableHtml += `<tr>`;
            headers.forEach(header => {
                tableHtml += `<td>${row[header] || ''}</td>`;
            });
            tableHtml += `</tr>`;
        });
        
        tableHtml += `</tbody></table>`;
        
        // 如果数据超过5行，显示提示信息
        if (data.length > 5) {
            tableHtml += `<p>...还有${data.length - 5}行未显示</p>`;
        }
        
        previewElement.innerHTML = tableHtml; // 更新预览元素内容
    }
    
    /**
     * 计算并显示统计数据
     * 优化：添加本期结余统计
     */
    function calculateAndDisplayStats() {
        if (summaryData.length > 0) {
            let consumptionTotal = 0; // 消费总额
            let depositTotal = 0; // 预存总额
            let settlementTotal = 0; // 本期结余总额
            
            // 遍历汇总数据计算总额
            summaryData.forEach(row => {
                const amount = row['金额（元）'] || 0; // 解析金额
                if (row['交易类型'] === '消费') {
                    consumptionTotal += amount; // 累加消费金额
                } else if (row['交易类型'] === '预存') {
                    depositTotal += amount; // 累加预存金额
                }
            });
            
            // 计算本期结余总额 = 预存总额 + 消费总额
            settlementTotal = depositTotal + consumptionTotal;
            
            // 更新统计显示
            consumptionStat.textContent = '¥' + consumptionTotal.toFixed(2);
            depositStat.textContent = '¥' + depositTotal.toFixed(2);
            balanceStat.textContent = '¥' + settlementTotal.toFixed(2); // 使用balanceStat显示本期结余
        }
    }
    
    // 处理数据按钮点击事件
    processBtn.addEventListener('click', async function() {
        processBtn.disabled = true; // 禁用按钮防止重复点击
        processBtn.textContent = '处理中...'; // 更新按钮文本
        progressArea.innerHTML = '正在处理数据，请稍候...'; // 更新进度信息
        
        try {
            // 模拟处理进度
            for (let i = 0; i <= 100; i += 20) {
                await new Promise(resolve => setTimeout(resolve, 300)); // 等待300毫秒
                progressBar.style.width = i + '%'; // 更新进度条宽度
                progressArea.innerHTML = `数据处理中: ${i}%`; // 更新进度文本
            }
            
            // 实际处理数据 - 根据上传的文件情况选择处理方式
            if (file1 && file2) {
                // 两个文件都上传，使用明细数据
                processedData = processSummaryData(summaryData, detailData);
            } else if (file1) {
                // 只上传了汇总表，使用空明细数据
                processedData = processSummaryData(summaryData, []);
            } else {
                throw new Error('请至少上传汇总表文件');
            }
            
            // 显示处理结果
            showDataPreview(processedData, resultPreview);
            progressArea.innerHTML = '数据处理完成！已生成完善的汇总表';
            resultSection.style.display = 'block'; // 显示结果区域
            processBtn.textContent = '处理数据并导出结果'; // 恢复按钮文本
            
        } catch (error) {
            console.error('数据处理错误:', error);
            progressArea.innerHTML = '处理数据时出错，请重试';
            progressArea.style.color = '#e74c3c'; // 设置错误颜色
            processBtn.disabled = false;
            processBtn.textContent = '处理数据并导出结果';
        }
    });
    
    /**
     * 创建余额汇总行
     * @param {Object} templateRow - 模板行（取第一个行作为模板）
     * @param {number} balance - 余额
     * @returns {Object} 余额汇总行
     */
    function createBalanceRow(templateRow, balance) {
        const balanceRow = {};
        for (const key in templateRow) {
            balanceRow[key] = templateRow[key];
        }
        balanceRow['交易类型'] = '余额汇总';
        balanceRow['金额（元）'] = balance;
        balanceRow['缴费日期'] = ''; // 清空缴费日期
        return balanceRow;
    }
    
    /**
     * 创建本期结余行
     * @param {Object} templateRow - 模板行（取第一个行作为模板）
     * @param {number} settlementAmount - 本期结余金额
     * @returns {Object} 本期结余行
     */
    function createSettlementRow(templateRow, settlementAmount) {
        const settlementRow = {};
        for (const key in templateRow) {
            settlementRow[key] = templateRow[key];
        }
        settlementRow['交易类型'] = '本期结余';
        settlementRow['金额（元）'] = settlementAmount;
        settlementRow['缴费日期'] = ''; // 清空缴费日期
        return settlementRow;
    }
    
    /**
     * 处理汇总数据，添加余额信息和本期结余
     * @param {Array} summaryData - 汇总表数据
     * @param {Array} detailData - 明细表数据
     * @returns {Array} 处理后的数据
     */
    function processSummaryData(summaryData, detailData) {
        const objectBalanceMap = new Map(); // 创建Map存储对象余额
        
        // 如果有明细数据，则提取每个对象编号的最新余额
        if (detailData.length > 0) {
            // 按对象编号分组明细数据
            const detailGrouped = {};
            detailData.forEach(row => {
                const objectId = row['对象编号'];
                if (!objectId) return;
                
                if (!detailGrouped[objectId]) {
                    detailGrouped[objectId] = [];
                }
                detailGrouped[objectId].push(row);
            });
            
        // 对每个对象的记录按创建时间降序排序，取最新记录
        Object.keys(detailGrouped).forEach(objectId => {
            const records = detailGrouped[objectId];
            records.sort((a, b) => {
                const timeA = a['创建时间'] ? new Date(a['创建时间']).getTime() : 0;
                const timeB = b['创建时间'] ? new Date(b['创建时间']).getTime() : 0;
                return timeB - timeA; // 降序排序
            });
            
            // 取第一条记录的交易后金额作为余额
            const latestRecord = records[0];
            if (latestRecord && latestRecord['交易后金额（元）'] !== undefined) {
                // 修复：确保余额是数字类型
                const balance = parseFloat(latestRecord['交易后金额（元）']);
                objectBalanceMap.set(objectId, isNaN(balance) ? 0 : balance);
            }
        });
        }
        
        // 按对象编号分组汇总数据
        const summaryGrouped = {};
        summaryData.forEach(row => {
            const objectId = row['对象编号'];
            if (!objectId) return;
            
            if (!summaryGrouped[objectId]) {
                summaryGrouped[objectId] = [];
            }
            summaryGrouped[objectId].push(row);
        });
        
        // 为每个对象创建余额汇总行和本期结余行
        const resultData = [];
        Object.keys(summaryGrouped).forEach(objectId => {
            const groupRows = summaryGrouped[objectId];
            
            // 添加该组的所有原始行
            groupRows.forEach(row => {
                // 修复：确保金额是数字类型
                if (typeof row['金额（元）'] !== 'number') {
                    row['金额（元）'] = parseFloat(row['金额（元）']) || 0;
                }
                resultData.push(row);
            });
            
            // 计算该用户的消费总额和预存总额
            let consumptionTotal = 0;
            let depositTotal = 0;
            
            groupRows.forEach(row => {
                const amount = row['金额（元）'] || 0;
                if (row['交易类型'] === '消费') {
                    consumptionTotal += amount;
                } else if (row['交易类型'] === '预存') {
                    depositTotal += amount;
                }
            });
            
            // 计算本期结余 = 预存总额 + 消费总额
            // 注意：消费通常是负数，所以实际上是预存减去消费的绝对值
            const settlementAmount = depositTotal + consumptionTotal;
            
            // 如果有明细数据，则创建余额汇总行
            if (detailData.length > 0) {
                const balance = objectBalanceMap.get(objectId) || 0;
                const balanceRow = createBalanceRow(groupRows[0], balance);
                resultData.push(balanceRow);
            }
            
            // 创建本期结余行
            const settlementRow = createSettlementRow(groupRows[0], settlementAmount);
            resultData.push(settlementRow);
        });
        
        // 计算本期结余总和
        let settlementTotal = 0;
        Object.keys(summaryGrouped).forEach(objectId => {
            const groupRows = summaryGrouped[objectId];
            let consumptionTotal = 0;
            let depositTotal = 0;
            
            groupRows.forEach(row => {
                const amount = row['金额（元）'] || 0;
                if (row['交易类型'] === '消费') {
                    consumptionTotal += amount;
                } else if (row['交易类型'] === '预存') {
                    depositTotal += amount;
                }
            });
            
            settlementTotal += depositTotal + consumptionTotal;
        });
        
        balanceStat.textContent = '¥' + settlementTotal.toFixed(2); // 更新余额显示
        
        return resultData; // 返回处理后的数据
    }
    
    // 下载按钮点击事件处理函数
    downloadBtn.addEventListener('click', function() {
        // 检查是否已处理数据
        if (!processedData) {
            alert('请先处理数据');
            return;
        }
        
        // 更新进度显示
        progressArea.innerHTML = '正在生成Excel文件...';
        
        try {
            // 创建工作表 - 将处理后的JSON数据转换为Excel工作表
            const ws = XLSX.utils.json_to_sheet(processedData);
            
            // 创建工作簿 - 创建新的Excel工作簿
            const wb = XLSX.utils.book_new();
            // 将工作表添加到工作簿中，并命名为"完整汇总表"
            XLSX.utils.book_append_sheet(wb, ws, "完整汇总表");
            
            // ========== 简化的标题行样式设置 ==========
            
            // 定义标题行样式对象
            const titleStyle = {
                // 字体样式设置
                font: {
                    name: '宋体',      // 字体为宋体
                    sz: 12,           // 字号为12
                    bold: true,       // 加粗显示
                    color: { rgb: "000000" } // 字体颜色为黑色
                },
                // 对齐方式设置
                alignment: {
                    horizontal: "center", // 水平居中
                    vertical: "center"   // 垂直居中
                },
                // 边框设置 - 四边都添加细黑边框
                border: {
                    top: { style: "thin", color: { rgb: "000000" } },
                    left: { style: "thin", color: { rgb: "000000" } },
                    bottom: { style: "thin", color: { rgb: "000000" } },
                    right: { style: "thin", color: { rgb: "000000" } }
                },
                // 填充颜色设置 - 浅蓝色背景
                fill: {
                    fgColor: { rgb: "DDEBF7" } // 背景色为浅蓝色
                }
            };
            
            // 直接为A到J列的第一行设置样式（列索引0-9对应A-J）
            for (let col = 0; col <= 9; col++) {
                // 将列索引和行号(0,0)转换为单元格地址（如A1、B1等）
                const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
                
                // 确保单元格存在，如果不存在则创建
                if (!ws[cellAddress]) {
                    ws[cellAddress] = { v: '' };
                }
                
                // 应用定义好的标题样式到当前单元格
                ws[cellAddress].s = titleStyle;
            }
            
            // 设置列宽 - 为A到J列设置统一的列宽
            const colWidths = [];
            for (let i = 0; i <= 9; i++) {
                colWidths.push({ wch: 15 }); // 每列宽度设置为15个字符
            }
            ws['!cols'] = colWidths; // 将列宽设置应用到工作表
            
            // 生成Excel文件并下载 - 使用指定文件名保存
            XLSX.writeFile(wb, '完整消费预存余额汇总表.xlsx');
            
            // 更新完成提示
            progressArea.innerHTML = '文件已生成并开始下载！';
            
        } catch (error) {
            // 错误处理 - 捕获并显示生成过程中的错误
            console.error('生成Excel文件时出错:', error);
            progressArea.innerHTML = '生成文件时出错: ' + error.message;
            progressArea.style.color = '#e74c3c'; // 错误信息显示为红色
        }
    });
});