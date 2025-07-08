// 检查XLSX库是否已成功加载
window.addEventListener('DOMContentLoaded', function() {
    // 检查XLSX库是否已成功加载
    if (typeof XLSX === 'undefined') {
        console.error('错误: XLSX库未成功加载，将尝试重新加载');
        // 尝试重新加载XLSX库
        loadXLSXLibrary();
    } else {
        console.log('XLSX库已成功加载');
    }

    // 初始化新UI交互
    initNewUIInteractions();
});

// 加载XLSX库的函数
function loadXLSXLibrary() {
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    
    // 设置加载成功回调
    script.onload = function() {
        console.log('XLSX库已成功重新加载');
        document.getElementById('status').textContent = 'XLSX库已成功加载，可以开始使用工具';
        document.getElementById('status').style.color = 'green';
    };
    
    // 设置加载失败回调
    script.onerror = function() {
        console.error('XLSX库加载失败');
        document.getElementById('status').innerHTML = 
            '无法加载Excel处理库(XLSX)，请检查网络连接或<a href="https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js" target="_blank">点击此处</a>手动下载后放在项目文件夹中';
        document.getElementById('status').style.color = 'red';
    };
    
    document.head.appendChild(script);
}

// 初始化新UI交互
function initNewUIInteractions() {
    // 功能选项卡切换
    const featureItems = document.querySelectorAll('.feature-item');
    const navTabs = document.querySelectorAll('.nav-tab');
    
    // 功能选择交互
    featureItems.forEach(item => {
        item.addEventListener('click', function() {
            const feature = this.getAttribute('data-feature');
            
            // 移除所有激活状态
            featureItems.forEach(i => i.classList.remove('active'));
            
            // 添加当前激活状态
            this.classList.add('active');
            
            // 如果是比对或统计功能，切换到对应页面
            if (feature === 'compare') {
                switchToPage('compare-page');
            } else if (feature === 'stats') {
                switchToPage('stats-page');
            }
            
            console.log('选择功能:', feature);
        });
    });
    
    // 导航标签切换
    navTabs.forEach(tab => {
        tab.addEventListener('click', function() {
            const page = this.getAttribute('data-page');
            
            // 移除所有激活状态
            navTabs.forEach(t => t.classList.remove('active'));
            
            // 添加当前激活状态
            this.classList.add('active');
            
            // 切换页面
            switchToPage(page);
        });
    });
    
    // 快速操作按钮
    const quickBtns = document.querySelectorAll('.quick-btn');
    quickBtns.forEach(btn => {
        btn.addEventListener('click', function() {
            const action = this.textContent.trim();
            console.log('快速操作:', action);
            
            // 添加点击效果
            this.style.transform = 'scale(0.95)';
            setTimeout(() => {
                this.style.transform = 'scale(1)';
            }, 150);
        });
    });
    
    // 初始化原有的导航功能
    initNavigation();
}

// 切换页面函数
function switchToPage(pageId) {
    // 获取所有页面
    const pages = document.querySelectorAll('.app-page');
    
    // 隐藏所有页面
    pages.forEach(page => {
        page.classList.remove('active');
    });
    
    // 显示目标页面
    const targetPage = document.getElementById(pageId);
    if (targetPage) {
        targetPage.classList.add('active');
    }
    
    // 更新导航标签状态
    const navTabs = document.querySelectorAll('.nav-tab');
    navTabs.forEach(tab => {
        if (tab.getAttribute('data-page') === pageId) {
            tab.classList.add('active');
        } else {
            tab.classList.remove('active');
        }
    });
    
    // 更新功能选项状态
    const featureItems = document.querySelectorAll('.feature-item');
    if (pageId === 'compare-page') {
        featureItems.forEach(item => {
            if (item.getAttribute('data-feature') === 'compare') {
                item.classList.add('active');
            } else {
                item.classList.remove('active');
            }
        });
    } else if (pageId === 'stats-page') {
        featureItems.forEach(item => {
            if (item.getAttribute('data-feature') === 'stats') {
                item.classList.add('active');
            } else {
                item.classList.remove('active');
            }
        });
    }
}

// 获取页面元素
const fileA_input = document.getElementById('fileA');
const fileB_input = document.getElementById('fileB');
const compareBtn = document.getElementById('compareBtn');
const statusDiv = document.getElementById('status');
const summaryDiv = document.getElementById('summary');
const summaryContent = document.getElementById('summaryContent');
const differencesContainer = document.getElementById('differences-container');
const differencesTableContainer = document.getElementById('differences-table-container');
const confirmUpdateBtn = document.getElementById('confirmUpdateBtn');
const undoUpdateBtn = document.getElementById('undoUpdateBtn'); // 撤销按钮
const quickUpdateBtn = document.getElementById('quick-update-btn'); // 快速更新按钮

// 姓名列固定使用"姓名"字段
const NAME_COLUMN = "姓名";

// 工作表选择区域元素
const sheetsA_container = document.getElementById('sheetsA-container');
const sheetsB_container = document.getElementById('sheetsB-container');
const sheetsA_list = document.getElementById('sheetsA-list');
const sheetsB_list = document.getElementById('sheetsB-list');
const selectAllA_btn = document.getElementById('selectAllA');
const selectAllB_btn = document.getElementById('selectAllB');

// 存储工作表信息
const sheetsInfoA = {
    loaded: false,
    sheetNames: [],
    selectedSheets: new Set()
};

const sheetsInfoB = {
    loaded: false,
    sheetNames: [],
    selectedSheets: new Set()
};

// 存储比对结果和更新数据
let globalComparisonResult = {
    differences: [],
    originalDataA: null,
    originalDataB: null,
    updatedDataA: null,
    keyColumn: '',
    readyToExport: false,
    lastUpdateState: null, // 用于存储上一次更新前的状态，实现撤销功能
    indexes: {             // 用于存储索引，提高查询性能
        dataA: null,
        dataB: null
    }
};

// 绑定按钮点击事件 - 移除旧监听器以避免潜在的重复绑定问题
compareBtn.removeEventListener('click', handleCompare);
compareBtn.addEventListener('click', function(e) {
    console.log('比对按钮被点击');
    
    // 添加点击效果
    this.innerHTML = '🔄 正在比对...';
    this.disabled = true;
    
    try {
        handleCompare().finally(() => {
            // 恢复按钮状态
            this.innerHTML = '🔍 立即比对';
            this.disabled = false;
        });
    } catch (error) {
        console.error('比对过程中发生错误:', error);
        statusDiv.textContent = `比对过程出错: ${error.message}`;
        statusDiv.style.color = 'red';
        showProgressBar(false); // 确保出错时隐藏进度条
        
        // 恢复按钮状态
        this.innerHTML = '🔍 立即比对';
        this.disabled = false;
    }
});

confirmUpdateBtn.removeEventListener('click', handleConfirmUpdate);
confirmUpdateBtn.addEventListener('click', handleConfirmUpdate);

undoUpdateBtn.removeEventListener('click', handleUndoUpdate);
undoUpdateBtn.addEventListener('click', handleUndoUpdate); // 添加撤销按钮事件监听

// 绑定文件上传事件
fileA_input.removeEventListener('change', handleFileUpload);
fileB_input.removeEventListener('change', handleFileUpload);
fileA_input.addEventListener('change', (e) => handleFileUpload(e, 'A'));
fileB_input.addEventListener('change', (e) => handleFileUpload(e, 'B'));

// 绑定全选按钮事件，使用函数阻止事件冒泡
selectAllA_btn.removeEventListener('click', null);
selectAllB_btn.removeEventListener('click', null);
selectAllA_btn.addEventListener('click', function(e) {
    e.stopPropagation();  // 阻止事件冒泡
    e.preventDefault();   // 阻止默认行为
    toggleAllSheets('A');
});
selectAllB_btn.addEventListener('click', function(e) {
    e.stopPropagation();  // 阻止事件冒泡
    e.preventDefault();   // 阻止默认行为
    toggleAllSheets('B');
});

/**
 * 查找两个数据集中所有需要更新的记录
 * @param {Array<Object>} dataA - 表格A的数据（旧数据）
 * @param {Array<Object>} dataB - 表格B的数据（新数据）
 * @returns {Array<Object>} - 需要更新的记录数组，每条记录包含旧值和新值
 */
/**
 * 异步查找两个数据集中所有需要更新的记录，使用分批处理避免UI阻塞
 * @param {Array<Object>} dataA - 表格A的数据（旧数据）
 * @param {Array<Object>} dataB - 表格B的数据（新数据）
 * @param {Function} progressCallback - 进度回调函数
 * @returns {Promise<Object>} - 包含需要更新和新增的记录数组的Promise
 */
function findDifferences(dataA, dataB, progressCallback) {
    return new Promise(async (resolve) => {
        const differences = [];  // 需要更新的记录
        const newStudents = [];  // 新增学生记录
        
        // 创建数据索引，提高查询效率
        if (progressCallback) progressCallback('正在创建数据索引...', 55);
        
        // 为B表创建索引
        globalComparisonResult.indexes.dataB = createMultiIndex(dataB);
        
        // 将B表数据转换为Map，便于快速查找，使用姓名作为键
        const dataB_Map = new Map();
        
        // 显示进度：准备B表数据
        if (progressCallback) progressCallback('正在索引新表数据...', 58);
        
        dataB.forEach(rowB => {
            if (rowB[NAME_COLUMN]) {
                const name = String(rowB[NAME_COLUMN]).trim();
                
                // 如果有重名学生，保留在Map中
                if (dataB_Map.has(name)) {
                    const existingData = dataB_Map.get(name);
                    if (Array.isArray(existingData)) {
                        existingData.push(rowB);
                    } else {
                        dataB_Map.set(name, [existingData, rowB]);
                    }
                } else {
                    dataB_Map.set(name, rowB);
                }
            }
        });
        
        // 为A表创建索引
        globalComparisonResult.indexes.dataA = createMultiIndex(dataA);
        
        // 创建A表姓名集合，用于后续判断新学生
        const namesInA = new Set();
        dataA.forEach(rowA => {
            if (rowA[NAME_COLUMN]) {
                namesInA.add(String(rowA[NAME_COLUMN]).trim());
            }
        });
        
        // 遍历A表数据，查找在B表中有更新的记录
        if (progressCallback) progressCallback('正在比对学生数据...', 60);
        
        // 分批处理参数
        const totalRows = dataA.length;
        const batchSize = 200; // 每批处理200条记录
        let processedRows = 0;
        
        // 分批处理数据
        for (let i = 0; i < totalRows; i += batchSize) {
            // 创建当前批次的处理范围
            const endIndex = Math.min(i + batchSize, totalRows);
            const currentBatch = dataA.slice(i, endIndex);
            
            // 使用setTimeout让UI有机会更新
            await new Promise(resolve => setTimeout(resolve, 0));
            
            // 处理当前批次
            for (const rowA of currentBatch) {
                // 更新比对进度
                processedRows++;
                if (progressCallback && processedRows % 50 === 0) { // 每50条记录更新一次进度
                    const progress = 60 + Math.floor((processedRows / totalRows) * 15);
                    progressCallback(`正在比对第 ${processedRows}/${totalRows} 条记录...`, progress);
                }
                
                if (!rowA[NAME_COLUMN]) continue; // 跳过没有姓名的记录
                
                const name = String(rowA[NAME_COLUMN]).trim();
                if (dataB_Map.has(name)) {
                    const rowBData = dataB_Map.get(name);
                    
                    // 处理重名情况
                    let rowB;
                    if (Array.isArray(rowBData)) {
                        // 如果有多个同名学生，选择最匹配的一个
                        rowB = findBestMatch(rowA, rowBData);
                    } else {
                        rowB = rowBData;
                    }
                    
                    const fieldDifferences = [];
                    
                    // 比较所有字段，查找差异
                    const allFields = new Set([
                        ...Object.keys(rowA),
                        ...Object.keys(rowB)
                    ]);
                    
                    let hasDifference = false;
                    
                    for (const field of allFields) {
                        // 跳过内部字段和姓名字段
                        if (field === '_原始工作表' || field === NAME_COLUMN || field.startsWith('_')) continue;
                        
                        const valueA = rowA[field];
                        const valueB = rowB[field];
                        
                        // 规范化比较值
                        const normalizedA = valueA !== null && valueA !== undefined ? String(valueA).trim() : '';
                        const normalizedB = valueB !== null && valueB !== undefined ? String(valueB).trim() : '';
                        
                        if (normalizedA !== normalizedB) {
                            hasDifference = true;
                            fieldDifferences.push({
                                field,
                                oldValue: valueA,
                                newValue: valueB
                            });
                        }
                    }
                    
                    // 如果存在差异，添加到结果中
                    if (hasDifference) {
                        differences.push({
                            key: name, // 使用姓名作为key
                            record: rowA['学号'] || rowA['考试号'] || '', // 尝试获取ID
                            name: name,
                            class: rowA['班级'] || rowA['班号'] || '', // 尝试获取班级
                            rowA,
                            rowB,
                            differences: fieldDifferences,
                            type: 'update' // 标记为更新类型
                        });
                    }
                }
            }
        }
        
        // 第二阶段：查找B表中存在但A表中不存在的学生（新学生）
        if (progressCallback) progressCallback('正在查找新增学生...', 75);
        
        // 遍历B表找出新学生
        let newStudentCount = 0;
        
        // 分批处理B表数据
        const totalRowsB = dataB.length;
        const batchSizeB = 200;
        
        for (let i = 0; i < totalRowsB; i += batchSizeB) {
            // 创建当前批次的处理范围
            const endIndex = Math.min(i + batchSizeB, totalRowsB);
            const currentBatch = dataB.slice(i, endIndex);
            
            // 使用setTimeout让UI有机会更新
            await new Promise(resolve => setTimeout(resolve, 0));
            
            // 处理当前批次
            for (const rowB of currentBatch) {
                if (!rowB[NAME_COLUMN]) continue; // 跳过没有姓名的记录
                
                const name = String(rowB[NAME_COLUMN]).trim();
                
                // 如果A表中不存在该学生
                if (!namesInA.has(name)) {
                    newStudentCount++;
                    
                    // 尝试获取班级和学号信息
                    const classField = rowB['班级'] || rowB['班号'] || rowB['班'] || '';
                    const studentId = rowB['学号'] || rowB['考试号'] || '';
                    
                    newStudents.push({
                        key: name + '_new', // 使用姓名+'_new'作为key，避免与更新记录冲突
                        record: studentId,
                        name: name,
                        class: classField,
                        rowA: null, // A表中不存在该记录
                        rowB: rowB, // B表中的记录
                        isNewStudent: true, // 标记为新学生
                        type: 'new' // 标记为新增类型
                    });
                    
                    if (progressCallback && newStudentCount % 20 === 0) {
                        progressCallback(`已找到 ${newStudentCount} 名新学生...`, 75);
                    }
                }
            }
        }
        
        // 完成比对
        if (progressCallback) {
            const totalDiffs = differences.length + newStudents.length;
            
            // 如果发现的差异很多，可能需要分析一下
            if (totalDiffs > 100) {
                progressCallback(`比对完成，发现 ${differences.length} 条需要更新的记录和 ${newStudents.length} 名新学生，正在分析...`, 80);
                // 添加分析统计数据
                const stats = analyzeUpdateDifferences(differences);
                console.log('差异分析:', stats);
            } else {
                progressCallback(`比对完成，共找到 ${differences.length} 条需要更新的记录和 ${newStudents.length} 名新学生`, 80);
            }
        }
        
        // 返回两种类型的结果
        resolve({
            differences: differences,
            newStudents: newStudents
        });
    });
}

/**
 * 分析差异数据
 * @param {Array<Object>} differences - 差异数据
 * @returns {Object} - 分析结果
 */
function analyzeUpdateDifferences(differences) {
    const stats = {
        totalUpdates: differences.length,
        byField: {},
        byClass: {},
        commonChanges: []
    };
    
    // 统计每个字段的更新次数
    differences.forEach(diff => {
        diff.differences.forEach(fieldDiff => {
            if (!stats.byField[fieldDiff.field]) {
                stats.byField[fieldDiff.field] = {
                    count: 0,
                    examples: []
                };
            }
            
            stats.byField[fieldDiff.field].count++;
            
            // 记录一些示例值，以便分析
            if (stats.byField[fieldDiff.field].examples.length < 5) {
                stats.byField[fieldDiff.field].examples.push({
                    name: diff.name,
                    oldValue: fieldDiff.oldValue,
                    newValue: fieldDiff.newValue
                });
            }
        });
        
        // 按班级统计
        const className = diff.class || '未知班级';
        if (!stats.byClass[className]) {
            stats.byClass[className] = 0;
        }
        stats.byClass[className]++;
    });
    
    // 识别常见的更改模式
    // 例如，如果大多数记录都更改了同一个字段，可能是系统性的更改
    for (const field in stats.byField) {
        if (stats.byField[field].count > differences.length * 0.3) { // 超过30%的记录更改了该字段
            stats.commonChanges.push({
                field,
                count: stats.byField[field].count,
                percentage: Math.round(stats.byField[field].count / differences.length * 100),
                examples: stats.byField[field].examples
            });
        }
    }
    
    return stats;
}

/**
 * 创建多重索引加速查找
 * @param {Array<Object>} data - 要索引的数据
 * @returns {Object} - 包含多个索引的对象
 */
function createMultiIndex(data) {
    // 创建索引对象
    const indexes = {
        byName: new Map(),      // 按姓名索引
        byClass: new Map(),     // 按班级索引
        byStudentId: new Map(), // 按学号索引
        byIdNumber: new Map()   // 按身份证号索引
    };
    
    // 遍历数据，创建索引
    data.forEach((row, index) => {
        // 按姓名索引
        if (row[NAME_COLUMN]) {
            const name = String(row[NAME_COLUMN]).trim();
            if (indexes.byName.has(name)) {
                const existing = indexes.byName.get(name);
                if (Array.isArray(existing)) {
                    existing.push(index);
                } else {
                    indexes.byName.set(name, [existing, index]);
                }
            } else {
                indexes.byName.set(name, index);
            }
        }
        
        // 按班级索引
        const classFields = ['班级', '班号', '班'];
        for (const field of classFields) {
            if (row[field]) {
                const className = String(row[field]).trim();
                if (!indexes.byClass.has(className)) {
                    indexes.byClass.set(className, []);
                }
                indexes.byClass.get(className).push(index);
                break;
            }
        }
        
        // 按学号索引
        const idFields = ['学号', '考试号', '准考证号'];
        for (const field of idFields) {
            if (row[field]) {
                const id = String(row[field]).trim();
                indexes.byStudentId.set(id, index);
                break;
            }
        }
        
        // 按身份证号索引
        if (row['身份证号']) {
            const idNumber = String(row['身份证号']).trim();
            indexes.byIdNumber.set(idNumber, index);
        }
    });
    
    return indexes;
}

/**
 * 从多个同名记录中找到最匹配的一个，使用加权评分系统
 * @param {Object} rowA - A表中的记录
 * @param {Array<Object>} rowBArray - B表中同名的多条记录
 * @returns {Object} - 最匹配的记录
 */
function findBestMatch(rowA, rowBArray) {
    // 定义匹配权重
    const weights = {
        '身份证号': 100,  // 身份证号是唯一标识，权重最高
        '学号': 80,      // 学号通常是唯一的
        '准考证号': 80,   // 准考证号通常是唯一的
        '考试号': 70,     // 考试号可能会变
        '班级': 60,       // 班级很重要，但可能会变
        '班号': 60,
        '班': 60,
        '性别': 20,       // 性别辅助匹配
        '出生日期': 50,    // 出生日期很有用
        '电话': 40,       // 电话有一定辨识度
        '年级': 30        // 年级有辅助作用
    };
    
    // 对每条记录进行评分
    const scoredMatches = rowBArray.map(rowB => {
        let score = 0;
        let matchDetails = [];
        
        // 遍历所有字段计算匹配度
        for (const field in rowA) {
            // 跳过内部字段
            if (field.startsWith('_') || !rowA[field] || !rowB[field]) continue;
            
            const valueA = String(rowA[field]).trim();
            const valueB = String(rowB[field]).trim();
            
            // 如果值完全相同
            if (valueA === valueB) {
                const fieldWeight = weights[field] || 10; // 默认权重为10
                score += fieldWeight;
                matchDetails.push(`${field}(+${fieldWeight})`);
            }
            // 对于部分字段，支持部分匹配
            else if (['姓名', '身份证号', '电话'].includes(field)) {
                // 检查部分匹配（例如身份证号后几位）
                if (valueA.length > 5 && valueB.length > 5) {
                    // 检查后4位是否匹配
                    const lastFourA = valueA.slice(-4);
                    const lastFourB = valueB.slice(-4);
                    if (lastFourA === lastFourB) {
                        const partialWeight = Math.floor(weights[field] * 0.7) || 7;
                        score += partialWeight;
                        matchDetails.push(`${field}部分匹配(+${partialWeight})`);
                    }
                }
            }
        }
        
        return {
            row: rowB,
            score,
            matchDetails
        };
    });
    
    // 按分数排序，取得分最高的记录
    scoredMatches.sort((a, b) => b.score - a.score);
    
    // 记录匹配详情到控制台，便于调试
    if (scoredMatches.length > 1) {
        console.log(`多个同名记录"${rowA[NAME_COLUMN]}"的匹配分数:`, 
            scoredMatches.map(m => `${m.score}分 (${m.matchDetails.join(', ')})`));
    }
    
    return scoredMatches[0].row;
}

/**
 * 更新数据集中的记录
 * @param {Array<Object>} dataA - 要更新的数据集（旧数据）
 * @param {Array<Object>} differences - 包含差异信息的记录数组
 * @returns {Array<Object>} - 更新后的数据集
 */
function updateRecords(dataA, diffData, progressCallback) {
    return new Promise(async (resolve) => {
        if (progressCallback) progressCallback('准备更新数据...', 85);
        
        // 处理传入参数，兼容新旧格式
        let differences = [];
        let newStudents = [];
        
        // 获取选中的差异
        const checkedDifferences = [];
        const checkedNewStudents = [];
        
        if (Array.isArray(diffData)) {
            // 旧格式：直接传入差异数组
            differences = diffData;
            
            // 获取选中的差异
            document.querySelectorAll('.diff-checkbox:checked').forEach(checkbox => {
                if (checkbox.dataset.diffType === 'update' || !checkbox.dataset.diffType) {
                    const diffIndex = parseInt(checkbox.dataset.diffIndex);
                    if (!isNaN(diffIndex) && differences[diffIndex]) {
                        checkedDifferences.push(differences[diffIndex]);
                    }
                }
            });
        } else if (diffData && typeof diffData === 'object') {
            // 新格式：包含差异和新学生的对象
            differences = diffData.differences || [];
            newStudents = diffData.newStudents || [];
            
            // 获取选中的差异和新学生
            document.querySelectorAll('.diff-checkbox:checked').forEach(checkbox => {
                if (checkbox.dataset.diffType === 'update' || !checkbox.dataset.diffType) {
                    const diffIndex = parseInt(checkbox.dataset.diffIndex);
                    if (!isNaN(diffIndex) && differences[diffIndex]) {
                        checkedDifferences.push(differences[diffIndex]);
                    }
                } else if (checkbox.dataset.diffType === 'new') {
                    const newStudentIndex = parseInt(checkbox.dataset.newStudentIndex);
                    if (!isNaN(newStudentIndex) && newStudents[newStudentIndex]) {
                        checkedNewStudents.push(newStudents[newStudentIndex]);
                    }
                }
            });
        }
        
        // 如果没有选中任何差异和新学生，直接返回
        if (checkedDifferences.length === 0 && checkedNewStudents.length === 0) {
            if (progressCallback) progressCallback('没有选择需要更新的记录', 0);
            resolve(dataA);
            return;
        }
        
        // 备份当前状态以支持撤销
        globalComparisonResult.lastUpdateState = {
            originalDataA: structuredClone(globalComparisonResult.originalDataA),
            updatedDataA: globalComparisonResult.updatedDataA ? structuredClone(globalComparisonResult.updatedDataA) : null,
            differences: structuredClone(globalComparisonResult.differences || []),
            newStudents: structuredClone(globalComparisonResult.newStudents || [])
        };
        
        // 将差异信息转换为Map，便于快速查找
        const diffMap = new Map();
        checkedDifferences.forEach(diff => {
            diffMap.set(diff.key, diff);
        });
        
        // 计算进度用
        const totalRows = dataA.length;
        const batchSize = 500; // 每批处理500条记录
        let processedRows = 0;
        let updatedCount = 0;
        
        // 创建结果数组
        let updatedDataA = [];
        
        // 分批处理数据
        for (let i = 0; i < totalRows; i += batchSize) {
            // 创建当前批次的处理范围
            const endIndex = Math.min(i + batchSize, totalRows);
            const currentBatch = dataA.slice(i, endIndex);
            
            // 使用setTimeout让UI有机会更新
            await new Promise(resolve => setTimeout(resolve, 0));
            
            // 处理当前批次
            const processedBatch = currentBatch.map(rowA => {
                // 更新处理进度
                processedRows++;
                if (progressCallback && processedRows % 100 === 0) { // 每100条记录更新一次进度
                    const progress = 85 + Math.floor((processedRows / totalRows) * 8);
                    progressCallback(`正在更新数据，已处理 ${processedRows}/${totalRows} 条记录...`, progress);
                }
                
                if (!rowA[NAME_COLUMN]) return rowA; // 跳过没有姓名的记录
                
                const name = String(rowA[NAME_COLUMN]).trim();
                if (diffMap.has(name)) {
                    const diff = diffMap.get(name);
                    
                    // 检查这条记录是否与差异记录匹配
                    // 对于重名情况，我们需要确保是同一条记录
                    if (diff.rowA !== rowA) {
                        // 如果是重名记录，检查是否是同一班级或其他关键字段
                        const classFields = ['班级', '班号', '班'];
                        let isMatch = false;
                        
                        for (const field of classFields) {
                            if (rowA[field] && diff.rowA[field] && 
                                String(rowA[field]).trim() === String(diff.rowA[field]).trim()) {
                                isMatch = true;
                                break;
                            }
                        }
                        
                        if (!isMatch) return rowA; // 跳过不匹配的重名记录
                    }
                    
                    updatedCount++;
                    
                    // 创建一个新对象，合并B表中的数据
                    const updatedRow = {
                        ...rowA, // 保留A表的所有字段
                        '_原始工作表': rowA['_原始工作表'],
                        '_已更新': true, // 添加标记，表示此记录已被更新
                        '_更新时间': new Date().toISOString(), // 记录更新时间
                        '_原始记录': { ...rowA } // 保存原始记录以供后续使用
                    };
                    
                    // 更新差异字段
                    diff.differences.forEach(fieldDiff => {
                        updatedRow[fieldDiff.field] = fieldDiff.newValue;
                    });
                    
                    return updatedRow;
                }
                
                return rowA;
            });
            
            // 添加到结果中
            updatedDataA = updatedDataA.concat(processedBatch);
        }
        
        // 处理新学生
        if (checkedNewStudents.length > 0) {
            if (progressCallback) progressCallback(`正在添加 ${checkedNewStudents.length} 名新学生...`, 93);
            
            // 按班级分组
            const classMappings = {};
            
            // 首先识别所有可能的班级
            updatedDataA.forEach(row => {
                if (!row[NAME_COLUMN]) return;
                
                // 获取班级字段
                const classFields = ['班级', '班号', '班'];
                for (const field of classFields) {
                    if (row[field]) {
                        const className = String(row[field]).trim();
                        if (!classMappings[className]) {
                            classMappings[className] = [];
                        }
                        classMappings[className].push(row);
                        break;
                    }
                }
            });
            
            // 准备要添加的新学生
            const newStudentsToAdd = [];
            
            // 处理每个新学生
            for (const newStudent of checkedNewStudents) {
                const className = newStudent.class;
                const studentId = newStudent.record;
                
                // 创建新学生记录
                const newRow = {
                    ...newStudent.rowB, // 使用B表中的记录
                    '_原始工作表': newStudent.rowB['_原始工作表'] || '新增学生',
                    '_是新学生': true, // 标记为新学生
                    '_已更新': '新增', // 添加"新增"标记，替代TRUE
                    '_添加时间': new Date().toISOString() // 记录添加时间
                };
                
                newStudentsToAdd.push({
                    newRow,
                    className,
                    studentId
                });
            }
            
            // 如果需要，对新学生进行排序
            newStudentsToAdd.sort((a, b) => {
                // 先按班级排序
                if (a.className !== b.className) {
                    return a.className.localeCompare(b.className);
                }
                
                // 再按学号排序
                return a.studentId.localeCompare(b.studentId);
            });
            
            // 将新学生插入到适当的位置
            const finalDataA = [];
            
            // 找到班级末尾，插入新学生
            let currentClass = null;
            let classRows = [];
            
            // 按班级分组处理
            for (const row of updatedDataA) {
                const rowClass = row['班级'] || row['班号'] || row['班'] || '';
                
                if (rowClass !== currentClass) {
                    // 如果班级变了，先处理上一个班级的记录
                    if (classRows.length > 0) {
                        finalDataA.push(...classRows);
                        
                        // 检查是否有属于这个班级的新学生
                        const studentsForClass = newStudentsToAdd.filter(s => s.className === currentClass);
                        if (studentsForClass.length > 0) {
                            finalDataA.push(...studentsForClass.map(s => s.newRow));
                        }
                    }
                    
                    // 开始新的班级
                    currentClass = rowClass;
                    classRows = [row];
                } else {
                    // 继续当前班级
                    classRows.push(row);
                }
            }
            
            // 处理最后一个班级
            if (classRows.length > 0) {
                finalDataA.push(...classRows);
                
                // 检查是否有属于这个班级的新学生
                const studentsForClass = newStudentsToAdd.filter(s => s.className === currentClass);
                if (studentsForClass.length > 0) {
                    finalDataA.push(...studentsForClass.map(s => s.newRow));
                }
            }
            
            // 添加没有匹配班级的新学生到最后
            const unmatchedStudents = newStudentsToAdd.filter(s => !classMappings[s.className]);
            if (unmatchedStudents.length > 0) {
                finalDataA.push(...unmatchedStudents.map(s => s.newRow));
            }
            
            updatedDataA = finalDataA;
        }
        
        // 性能分析与日志
        console.log(`更新完成，处理了 ${totalRows} 条记录，更新了 ${updatedCount} 条记录，添加了 ${checkedNewStudents.length} 名新学生`);
        
        if (progressCallback) {
            const message = checkedNewStudents.length > 0 
                ? `数据更新完成，已更新 ${updatedCount} 条记录，添加 ${checkedNewStudents.length} 名新学生` 
                : `数据更新完成，已更新 ${updatedCount} 条记录`;
            progressCallback(message, 95);
        }
        
        resolve(updatedDataA);
    });
}

/**
 * 处理文件上传事件，显示文件中的工作表
 * @param {Event} event - 文件上传事件
 * @param {string} fileType - 文件类型标识('A' 或 'B')
 */
function handleFileUpload(event, fileType) {
    const file = event.target.files[0];
    if (!file) return;

    // 检查XLSX对象是否已加载
    if (typeof XLSX === 'undefined') {
        statusDiv.textContent = `读取表格${fileType}失败: XLSX库未加载，正在尝试重新加载...`;
        statusDiv.style.color = 'red';
        loadXLSXLibrary();
        return;
    }

    const sheetsInfo = fileType === 'A' ? sheetsInfoA : sheetsInfoB;
    const uploadArea = fileType === 'A' ? document.getElementById('upload-area-A') : document.getElementById('upload-area-B');
    const sheetsList = fileType === 'A' ? document.getElementById('sheetsA-list') : document.getElementById('sheetsB-list');
    const sheetsContainer = fileType === 'A' ? document.getElementById('sheetsA-container') : document.getElementById('sheetsB-container');
    
    // 更新上传区域状态
    uploadArea.classList.add('file-uploaded');
    
    // 显示文件名
    const fileInfo = uploadArea.querySelector('.file-info');
    const fileName = uploadArea.querySelector('.file-name');
    fileName.textContent = file.name;
    fileInfo.classList.remove('hidden');

    // 清空之前的工作表列表
    sheetsInfo.loaded = false;
    sheetsInfo.sheetNames = [];
    sheetsInfo.selectedSheets.clear();
    sheetsList.innerHTML = '';
    
    // 显示处理状态
    statusDiv.textContent = `正在读取${fileType === 'A' ? '表格A' : '表格B'}的工作表...`;
    statusDiv.style.color = 'blue';

    // 读取Excel文件中的工作表
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            // 再次检查XLSX对象是否可用
            if (typeof XLSX === 'undefined') {
                throw new Error('XLSX库未正确加载，请刷新页面后重试');
            }
            
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // 获取所有工作表名
            const sheetNames = workbook.SheetNames;
            sheetsInfo.sheetNames = sheetNames;
            
            // 显示工作表选择区域
            sheetsContainer.style.display = 'block';
            
            // 创建工作表复选框
            sheetNames.forEach(sheetName => {
                const sheetItem = document.createElement('div');
                sheetItem.className = 'sheet-item';
                
                // 添加点击事件阻止冒泡
                sheetItem.addEventListener('click', function(e) {
                    e.stopPropagation();
                });
                
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.id = `sheet-${fileType}-${sheetName}`;
                checkbox.dataset.sheet = sheetName;
                checkbox.dataset.fileType = fileType;  // 添加文件类型标识
                checkbox.checked = true; // 默认选中
                
                // 添加到已选择集合
                sheetsInfo.selectedSheets.add(sheetName);
                
                // 记录每个工作表的添加
                console.log(`工作表 ${sheetName} 已添加到选择集合，当前已选: ${sheetsInfo.selectedSheets.size}`);
                
                const label = document.createElement('label');
                label.htmlFor = checkbox.id;
                label.textContent = sheetName;
                
                // 监听复选框变化
                checkbox.addEventListener('change', (event) => {
                    // 阻止事件冒泡，避免干扰其他事件处理
                    event.stopPropagation();
                    
                    // 强制确保复选框状态与事件一致
                    const isChecked = event.target.checked;
                    const sheetName = event.target.dataset.sheet;
                    
                    if (!sheetName) {
                        console.warn(`警告：复选框缺少dataset.sheet属性`);
                        return;
                    }
                    
                    if (isChecked) {
                        sheetsInfo.selectedSheets.add(sheetName);
                        console.log(`工作表 ${sheetName} 被选中`);
                    } else {
                        sheetsInfo.selectedSheets.delete(sheetName);
                        console.log(`工作表 ${sheetName} 被取消选中`);
                    }
                    
                    // 重新统计选中的工作表
                    const selectedCount = Array.from(sheetsList.querySelectorAll('input[type="checkbox"]:checked')).length;
                    console.log(`当前有 ${selectedCount} 个工作表被选中，选择集合大小为 ${sheetsInfo.selectedSheets.size}`);
                    
                    // 更新已选工作表数量显示
                    updateSelectedSheetsCount(fileType);
                    
                    // 更新全选按钮文本
                    updateSelectAllButtonText(fileType);
                    
                    // 详细记录选择状态
                    const selectedSheets = Array.from(sheetsInfo.selectedSheets);
                    console.log(`选择状态更新: 文件${fileType}已选择 ${selectedSheets.length} 个工作表: ${selectedSheets.join(', ')}`);
                });
                
                sheetItem.appendChild(checkbox);
                sheetItem.appendChild(label);
                sheetsList.appendChild(sheetItem);
            });
            
            // 更新全选按钮文本
            updateSelectAllButtonText(fileType);
            
            // 更新已选工作表数量显示
            updateSelectedSheetsCount(fileType);
            
            sheetsInfo.loaded = true;
            
            statusDiv.textContent = `表格${fileType}已加载，共有 ${sheetNames.length} 个工作表`;
            statusDiv.style.color = 'green';
            
            // 记录日志
            console.log(`表格${fileType}已加载 ${sheetNames.length} 个工作表，默认全选`);
        } catch (error) {
            console.error('读取Excel文件出错:', error);
            statusDiv.textContent = `读取表格${fileType}失败: ${error.message}`;
            statusDiv.style.color = 'red';
        }
    };
    
    reader.onerror = () => {
        statusDiv.textContent = `读取表格${fileType}失败`;
        statusDiv.style.color = 'red';
    };
    
    reader.readAsArrayBuffer(file);
}

/**
 * 更新已选工作表数量显示
 * @param {string} fileType - 文件类型标识('A' 或 'B')
 */
function updateSelectedSheetsCount(fileType) {
    const sheetsInfo = fileType === 'A' ? sheetsInfoA : sheetsInfoB;
    const sheetsContainer = fileType === 'A' ? document.getElementById('sheetsA-container') : document.getElementById('sheetsB-container');
    
    const sheetsCount = sheetsContainer.querySelector('.sheets-count');
    if (sheetsCount) {
        const selectedCount = sheetsInfo.selectedSheets.size;
        const totalCount = sheetsInfo.sheetNames.length;
        
        sheetsCount.textContent = selectedCount;
    }
}

/**
 * 更新全选按钮文本
 * @param {string} fileType - 文件类型标识('A' 或 'B')
 */
function updateSelectAllButtonText(fileType) {
    const sheetsInfo = fileType === 'A' ? sheetsInfoA : sheetsInfoB;
    const selectAllBtn = fileType === 'A' ? selectAllA_btn : selectAllB_btn;
    const sheetsList = fileType === 'A' ? sheetsA_list : sheetsB_list;
    
    // 获取实际选中的工作表数量
    const checkedCount = sheetsList.querySelectorAll('input[type="checkbox"]:checked').length;
    const totalCount = sheetsInfo.sheetNames.length;
    
    // 更新按钮文本
    selectAllBtn.textContent = checkedCount === totalCount ? '取消全选' : '全选';
    
    // 记录日志
    console.log(`更新${fileType}表全选按钮: 已选择${checkedCount}/${totalCount}个工作表, 按钮文本='${selectAllBtn.textContent}'`);
}

/**
 * 切换全选/取消全选工作表
 * @param {string} fileType - 文件类型标识('A' 或 'B')
 */
function toggleAllSheets(fileType) {
    // 获取相关元素和状态
    const sheetsInfo = fileType === 'A' ? sheetsInfoA : sheetsInfoB;
    const sheetsList = fileType === 'A' ? sheetsA_list : sheetsB_list;
    const selectAllBtn = fileType === 'A' ? selectAllA_btn : selectAllB_btn;
    
    console.log(`执行${fileType}表全选/取消全选操作`);
    
    // 获取当前选中的复选框数量
    const checkboxes = Array.from(sheetsList.querySelectorAll('input[type="checkbox"]'));
    const checkedCount = checkboxes.filter(cb => cb.checked).length;
    const totalCount = checkboxes.length;
    
    // 判断是全选还是取消全选
    const shouldCheck = checkedCount < totalCount;
    
    console.log(`当前状态: ${checkedCount}/${totalCount} 个工作表被选中，将${shouldCheck ? '全选' : '取消全选'}所有工作表`);
    
    // 清空已选择集合，准备重新填充
    sheetsInfo.selectedSheets.clear();
    
    // 更新所有复选框状态
    checkboxes.forEach(checkbox => {
        // 更新复选框选中状态
        checkbox.checked = shouldCheck;
        
        // 如果是选中状态，将工作表名添加到选择集合
        if (shouldCheck && checkbox.dataset.sheet) {
            sheetsInfo.selectedSheets.add(checkbox.dataset.sheet);
        }
    });
    
    // 更新全选按钮文本
    selectAllBtn.textContent = shouldCheck ? '取消全选' : '全选';
    
    // 记录选择结果
    console.log(`${fileType}表${shouldCheck ? '全选' : '取消全选'}完成，当前已选: ${sheetsInfo.selectedSheets.size} 个工作表`);
    if (sheetsInfo.selectedSheets.size > 0) {
        console.log(`已选工作表: ${Array.from(sheetsInfo.selectedSheets).join(', ')}`);
    }
}

/**
 * 处理比对按钮点击事件
 */
async function handleCompare() {
    // 记录开始时间，用于性能分析
    const startTime = performance.now();
    
    // 隐藏之前的结果摘要和差异表格（如果有）
    if (summaryDiv) summaryDiv.classList.add('hidden');
    if (differencesContainer) differencesContainer.classList.add('hidden');
    
    // 重置撤销按钮状态
    if (undoUpdateBtn) {
        undoUpdateBtn.disabled = true;
    }
    
    // 检查XLSX对象是否已加载
    if (typeof XLSX === 'undefined') {
        statusDiv.textContent = '错误：Excel处理库(XLSX)未加载，正在尝试重新加载...';
        statusDiv.style.color = 'red';
        loadXLSXLibrary();
        return;
    }
    
    // 1. 获取输入文件
    const fileA = fileA_input.files[0];
    const fileB = fileB_input.files[0];

    // 2. 输入验证
    if (!fileA || !fileB) {
        statusDiv.textContent = '错误：请确保已上传两个Excel文件！';
        statusDiv.style.color = 'red';
        return;
    }
    
    // 验证工作表选择状态
    console.log('准备比对，验证工作表选择状态...');
    console.log('表格A工作表状态:', sheetsInfoA);
    console.log('表格B工作表状态:', sheetsInfoB);
    
    // 同步工作表选择状态 - 从DOM中获取实际选中的复选框状态
    sheetsInfoA.selectedSheets.clear();
    sheetsInfoB.selectedSheets.clear();
    
    // 获取A表选中的工作表
    const checkedSheetsA = document.querySelectorAll('#sheetsA-list input[type="checkbox"]:checked');
    console.log(`从DOM获取的A表选中工作表数量: ${checkedSheetsA.length}`);
    
    checkedSheetsA.forEach(checkbox => {
        if (checkbox.dataset.sheet) {
            sheetsInfoA.selectedSheets.add(checkbox.dataset.sheet);
            console.log(`添加A表选中工作表: ${checkbox.dataset.sheet}`);
        }
    });
    
    // 获取B表选中的工作表
    const checkedSheetsB = document.querySelectorAll('#sheetsB-list input[type="checkbox"]:checked');
    console.log(`从DOM获取的B表选中工作表数量: ${checkedSheetsB.length}`);
    
    checkedSheetsB.forEach(checkbox => {
        if (checkbox.dataset.sheet) {
            sheetsInfoB.selectedSheets.add(checkbox.dataset.sheet);
            console.log(`添加B表选中工作表: ${checkbox.dataset.sheet}`);
        }
    });

    // 检查是否选择了要比对的工作表
    if (sheetsInfoA.selectedSheets.size === 0 || sheetsInfoB.selectedSheets.size === 0) {
        statusDiv.textContent = '错误：请至少选择一个工作表进行比对！';
        statusDiv.style.color = 'red';
        console.error(`比对失败: A表选中${sheetsInfoA.selectedSheets.size}个工作表, B表选中${sheetsInfoB.selectedSheets.size}个工作表`);
        return;
    }
    
    // 记录当前选择的工作表
    console.log('比对前工作表选择状态:');
    console.log('表格A选择的工作表:', Array.from(sheetsInfoA.selectedSheets));
    console.log('表格B选择的工作表:', Array.from(sheetsInfoB.selectedSheets));
    
    // 显示处理状态和进度条
    showProgressBar(true);
    
    // 创建进度回调函数
    const updateProgress = (message, percent) => {
        const progressBar = document.querySelector('.progress-bar');
        if (progressBar) {
            statusDiv.textContent = message;
            progressBar.style.width = `${percent}%`;
        }
    };
    
    // 显示初始进度
    updateProgress('正在准备比对...', 5);
    
    // 添加日志以帮助调试
    console.log(`开始比对，主键列: "${NAME_COLUMN}"`);

    try {
        // 3. 读取Excel文件并按工作表分开处理
        updateProgress('正在读取Excel文件...', 10);
        
        // 3.1 读取A表所有选定工作表的数据
        const allSheetsDataA = {};
        const allDataA = []; // 仍然保留合并数据，用于总览
        
        // 读取A表所有工作表，但按工作表分别保存
        const workbookA = await readExcelFile(fileA);
        
        // 计算总工作表数量，用于进度计算
        const totalSheets = sheetsInfoA.selectedSheets.size + sheetsInfoB.selectedSheets.size;
        let processedSheets = 0;
        
        // 获取已选中的工作表数组
        const selectedSheetsA = Array.from(sheetsInfoA.selectedSheets);
        
        // 处理A表的每个选中工作表
        for (const sheetName of selectedSheetsA) {
            if (workbookA.SheetNames.includes(sheetName)) {
                const worksheet = workbookA.Sheets[sheetName];
                let sheetData = XLSX.utils.sheet_to_json(worksheet);
                
                // 为每条记录添加工作表标识
                sheetData = sheetData.map(row => ({
                    ...row,
                    '_原始工作表': sheetName
                }));
                
                // 保存到对应工作表
                allSheetsDataA[sheetName] = sheetData;
                
                // 同时添加到合并数据中
                allDataA.push(...sheetData);
                
                // 更新进度
                processedSheets++;
                updateProgress(`正在读取A表工作表: ${sheetName}`, 10 + Math.floor((processedSheets / totalSheets) * 20));
            }
        }
        
        // 3.2 读取B表所有选定工作表的数据
        const allSheetsDataB = {};
        const allDataB = []; // 仍然保留合并数据，用于总览
        
        // 读取B表所有工作表，但按工作表分别保存
        const workbookB = await readExcelFile(fileB);
        
        // 获取已选中的工作表数组
        const selectedSheetsB = Array.from(sheetsInfoB.selectedSheets);
        
        // 处理B表的每个选中工作表
        for (const sheetName of selectedSheetsB) {
            if (workbookB.SheetNames.includes(sheetName)) {
                const worksheet = workbookB.Sheets[sheetName];
                let sheetData = XLSX.utils.sheet_to_json(worksheet);
                
                // 为每条记录添加工作表标识
                sheetData = sheetData.map(row => ({
                    ...row,
                    '_原始工作表': sheetName
                }));
                
                // 保存到对应工作表
                allSheetsDataB[sheetName] = sheetData;
                
                // 同时添加到合并数据中
                allDataB.push(...sheetData);
                
                // 更新进度
                processedSheets++;
                updateProgress(`正在读取B表工作表: ${sheetName}`, 10 + Math.floor((processedSheets / totalSheets) * 20));
            }
        }
        
        // 检查是否成功读取到数据
        if (allDataA.length === 0 || allDataB.length === 0) {
            throw new Error(`读取工作表数据失败: A表数据${allDataA.length}条, B表数据${allDataB.length}条`);
        }
        
        console.log(`成功读取数据: A表${allDataA.length}条记录, B表${allDataB.length}条记录`);
        
        // 检查数据规模，提供性能优化提示
        if (allDataA.length > 5000 || allDataB.length > 5000) {
            console.warn(`大数据集警告: 表格A有${allDataA.length}行，表格B有${allDataB.length}行，可能需要较长处理时间`);
        }
        
        // 4. 使用Promise.race来设置超时，防止长时间阻塞UI
        const timeoutPromise = new Promise((_, reject) => {
            setTimeout(() => reject(new Error('比对操作超时，请检查数据规模或尝试刷新页面')), 120000); // 120秒超时
        });
        
        // 5. 执行按工作表分组的比对逻辑
        const comparePromise = new Promise(async (resolve) => {
            updateProgress('正在准备分工作表比对...', 30);
            
            // 按工作表分别处理比对
            const allResults = await compareDataBySheets(allSheetsDataA, allSheetsDataB, updateProgress);
            
            // 保存到全局变量
            globalComparisonResult.originalDataA = structuredClone(allDataA);
            globalComparisonResult.originalDataB = structuredClone(allDataB);
            globalComparisonResult.keyColumn = NAME_COLUMN;
            globalComparisonResult.differences = allResults.combined.differences;
            globalComparisonResult.newStudents = allResults.combined.newStudents;
            globalComparisonResult.bySheet = allResults.bySheet;
            globalComparisonResult.sheetStats = allResults.sheetStats;
            
            // 计算总差异数量
            const totalDifferences = allResults.combined.differences.length;
            const totalNewStudents = allResults.combined.newStudents.length;
            const totalChanges = totalDifferences + totalNewStudents;
            
            // 显示差异表格
            if (totalChanges > 0) {
                differencesContainer.classList.remove('hidden');
                displayDifferencesWithSheetTabs(allResults);
            } else {
                differencesContainer.classList.add('hidden');
            }
            
            // 生成结果摘要
            generateResultsSummaryWithSheets(allResults, allDataA.length, allDataB.length);
            
            resolve(allResults);
        });
        
        // 等待比对完成或超时
        const result = await Promise.race([comparePromise, timeoutPromise]);
        
        // 完成进度条
        updateProgress('比对完成', 100);
        
        // 计算耗时
        const endTime = performance.now();
        const timeSpent = ((endTime - startTime) / 1000).toFixed(2);
        
        // 计算总更改数
        const diffCount = globalComparisonResult.differences ? globalComparisonResult.differences.length : 0;
        const newStudentCount = globalComparisonResult.newStudents ? globalComparisonResult.newStudents.length : 0;
        const totalChanges = diffCount + newStudentCount;
        
        // 如果发现了差异，提示用户选择要更新的数据
        if (totalChanges > 0) {
            let message = `比对完成！耗时${timeSpent}秒，`;
            
            if (diffCount > 0 && newStudentCount > 0) {
                message += `发现 ${diffCount} 条需要更新的数据和 ${newStudentCount} 名新学生，`;
            } else if (diffCount > 0) {
                message += `发现 ${diffCount} 条需要更新的数据，`;
            } else if (newStudentCount > 0) {
                message += `发现 ${newStudentCount} 名新学生，`;
            }
            
            message += `可以使用工作表选项卡切换查看各工作表的结果，请选择要更新的记录，然后点击确认更新按钮。`;
            
            statusDiv.textContent = message;
            statusDiv.style.color = 'green';
            summaryDiv.classList.remove('hidden');
        } else {
            statusDiv.textContent = `比对完成！耗时${timeSpent}秒，两表数据完全一致，没有发现差异。`;
            statusDiv.style.color = 'green';
        }
        
        // 隐藏进度条
        setTimeout(() => {
            showProgressBar(false);
        }, 1500);

    } catch (error) {
        console.error('比对过程中出错:', error);
        updateProgress(`处理失败：${error.message}`, 0);
        statusDiv.style.color = 'red';
        
        // 隐藏进度条
        setTimeout(() => {
            showProgressBar(false);
        }, 3000);
    }
}

/**
 * 创建比对结果
 * @param {Array<Object>} dataA - 表格A的数据
 * @param {Array<Object>} dataB - 表格B的数据
 * @param {Array<Object>} differences - 差异数据
 * @param {Function} progressCallback - 进度回调函数
 * @returns {Promise<Array<Object>>} - 比对结果
 */
async function createComparisonResult(dataA, dataB, differences, progressCallback) {
    return new Promise(async (resolve) => {
        if (progressCallback) progressCallback('正在生成比对结果...', 85);
        
        // 将差异信息转换为Map，便于快速查找
        const diffMap = new Map();
        differences.forEach(diff => {
            diffMap.set(diff.key, diff);
        });
        
        // 批量处理
        const result = [];
        const batchSize = 1000; // 每批处理1000条记录
        
        // 处理表格A中的记录
        for (let i = 0; i < dataA.length; i += batchSize) {
            // 使用setTimeout让UI有机会更新
            await new Promise(resolve => setTimeout(resolve, 0));
            
            const endIdx = Math.min(i + batchSize, dataA.length);
            const percent = 85 + Math.floor((i / dataA.length) * 10);
            
            if (progressCallback) {
                progressCallback(`正在生成比对结果 ${i}/${dataA.length}...`, percent);
            }
            
            // 处理当前批次
            for (let j = i; j < endIdx; j++) {
                const rowA = dataA[j];
                
                if (!rowA[NAME_COLUMN]) continue;
                
                const name = String(rowA[NAME_COLUMN]).trim();
                
                if (diffMap.has(name)) {
                    // 有差异的记录
                    const diff = diffMap.get(name);
                    
                    result.push({
                        ...rowA,
                        '比对状态': '数据存在差异',
                        '差异说明': diff.differences.map(d => `${d.field}: 旧值="${d.oldValue}" 新值="${d.newValue}"`).join('; ')
                    });
                } else if (showIdentical.checked) {
                    // 相同的记录，如果选择了显示相同记录
                    result.push({
                        ...rowA,
                        '比对状态': '完全相同',
                        '差异说明': '两表中数据完全一致'
                    });
                }
            }
        }
        
        if (progressCallback) progressCallback('比对结果生成完成', 95);
        
        resolve(result);
    });
}

/**
 * 使用 SheetJS 解析 Excel 文件，返回 JSON 对象数组
 * @param {File} file - 用户上传的文件
 * @param {Set<string>} selectedSheets - 选择的工作表集合
 * @returns {Promise<Array<Object>>}
 */
function parseExcel(file, selectedSheets) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (event) => {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // 获取所有工作表名
                const sheetNames = workbook.SheetNames;
                
                // 用于存储所有工作表的合并数据
                let allData = [];
                
                // 记录选中的工作表
                const selectedSheetsList = Array.from(selectedSheets);
                console.log(`处理Excel文件 ${file.name}，处理选中的工作表: ${selectedSheetsList.join(', ')}`);
                
                // 调试信息
                console.log(`工作表选择情况: 总数=${sheetNames.length}, 选中=${selectedSheets.size}`);
                console.log(`所有工作表: ${sheetNames.join(', ')}`);
                console.log(`选中的工作表: ${selectedSheetsList.join(', ')}`);
                
                // 如果没有选中任何工作表，提供警告
                if (selectedSheets.size === 0) {
                    console.warn(`警告: 没有选中任何工作表进行处理, ${file.name} 的可用工作表有: ${sheetNames.join(', ')}`);
                }
                
                // 遍历工作表并合并数据
                for (const sheetName of sheetNames) {
                    // 检查是否选中
                    if (selectedSheets.has(sheetName)) {
                        const worksheet = workbook.Sheets[sheetName];
                        const sheetData = XLSX.utils.sheet_to_json(worksheet);
                        
                        // 记录日志
                        console.log(`读取工作表: ${sheetName}, 包含 ${sheetData.length} 条记录`);
                        
                        // 为每条记录添加工作表标识
                        const dataWithSheet = sheetData.map(row => ({
                            ...row,
                            '_原始工作表': sheetName
                        }));
                        
                        // 合并到总数据中
                        allData = allData.concat(dataWithSheet);
                    } else {
                        console.log(`跳过未选中的工作表: ${sheetName}`);
                    }
                }
                
                console.log(`处理了 ${selectedSheets.size} 个工作表，共 ${allData.length} 条记录`);
                
                // 验证数据完整性
                if (selectedSheets.size > 0 && allData.length === 0) {
                    console.warn(`警告: 所选工作表中没有找到数据，请检查表格格式`);
                }
                
                resolve(allData);
            } catch (e) {
                console.error('解析Excel文件失败:', e);
                reject(e);
            }
        };
        reader.onerror = (error) => reject(error);
        reader.readAsArrayBuffer(file);
    });
}

/**
 * 显示需要更新的数据差异和新学生
 * @param {Object} data - 包含差异记录和新学生的对象
 */
function displayDifferences(data) {
    // 处理传入参数，兼容新旧格式
    let differences = [];
    let newStudents = [];
    
    if (Array.isArray(data)) {
        // 旧格式：直接传入差异数组
        differences = data;
    } else if (data && typeof data === 'object') {
        // 新格式：包含差异和新学生的对象
        differences = data.differences || [];
        newStudents = data.newStudents || [];
    }
    
    const totalItems = differences.length + newStudents.length;
    
    // 清空容器
    differencesTableContainer.innerHTML = '';
    
    if (totalItems === 0) {
        differencesTableContainer.innerHTML = '<p>没有发现需要更新的数据</p>';
        return;
    }
    
    // 确保groupByClass参数不再被使用，移除相关功能
    if (groupByClass && groupByClass.checked) {
        groupByClass.checked = false;
    }
    
    // 获取A表和B表的记录数
    const totalA = globalComparisonResult.originalDataA ? globalComparisonResult.originalDataA.length : 0;
    const totalB = globalComparisonResult.originalDataB ? globalComparisonResult.originalDataB.length : 0;
    
    // 计算仅在A表存在的记录数（通过名称比较）
    const namesInB = new Set();
    if (globalComparisonResult.originalDataB) {
        globalComparisonResult.originalDataB.forEach(row => {
            if (row[NAME_COLUMN]) {
                namesInB.add(String(row[NAME_COLUMN]).trim());
            }
        });
    }
    
    let onlyInA = 0;
    if (globalComparisonResult.originalDataA) {
        globalComparisonResult.originalDataA.forEach(row => {
            if (row[NAME_COLUMN] && !namesInB.has(String(row[NAME_COLUMN]).trim())) {
                onlyInA++;
            }
        });
    }
    
    // 创建统计卡片
    const statsCards = document.getElementById('stats-cards');
    statsCards.innerHTML = '';
    
    // 创建统计卡片容器
    const statsContainer = document.createElement('div');
    statsContainer.className = 'stats-dashboard';
    
    // 计算仅在B表存在的记录数
    let onlyInB = 0;
    if (globalComparisonResult.originalDataB) {
        const namesInA = new Set();
        if (globalComparisonResult.originalDataA) {
            globalComparisonResult.originalDataA.forEach(row => {
                if (row[NAME_COLUMN]) {
                    namesInA.add(String(row[NAME_COLUMN]).trim());
                }
            });
        }
        
        globalComparisonResult.originalDataB.forEach(row => {
            if (row[NAME_COLUMN] && !namesInA.has(String(row[NAME_COLUMN]).trim())) {
                onlyInB++;
            }
        });
    }
    
    // 创建数据总览卡片
    const dataOverview = document.createElement('div');
    dataOverview.className = 'stats-overview';
    dataOverview.innerHTML = `
        <div class="stats-section">
            <div class="stats-header">
                <h3>数据比对总览</h3>
                <span class="stats-badge">共处理${totalA + totalB}条记录</span>
            </div>
            <div class="stats-cards-row">
                <div class="stat-card primary">
                    <div class="stat-card-header">
                        <div class="stat-card-title">表格A总记录</div>
                        <div class="stat-card-icon">
                            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <path d="M14 2H6a2 2 0 0 0-2 2v16c0 1.1.9 2 2 2h12a2 2 0 0 0 2-2V8l-6-6z"/>
                                <path d="M14 3v5h5M16 13H8M16 17H8M10 9H8"/>
                            </svg>
                        </div>
                    </div>
                    <div class="stat-card-value">${totalA}</div>
                    <div class="stat-card-footer">
                        <span>其中有${onlyInA}条仅在A表存在</span>
                    </div>
                </div>
                
                <div class="stat-card secondary">
                    <div class="stat-card-header">
                        <div class="stat-card-title">表格B总记录</div>
                        <div class="stat-card-icon">
                            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <path d="M14 2H6a2 2 0 0 0-2 2v16c0 1.1.9 2 2 2h12a2 2 0 0 0 2-2V8l-6-6z"/>
                                <path d="M14 3v5h5M16 13H8M16 17H8M10 9H8"/>
                            </svg>
                        </div>
                    </div>
                    <div class="stat-card-value">${totalB}</div>
                    <div class="stat-card-footer">
                        <span>其中有${onlyInB}条仅在B表存在</span>
                    </div>
                </div>
            </div>
        </div>
    `;
    statsContainer.appendChild(dataOverview);
    
    // 创建处理详情卡片
    const processDetails = document.createElement('div');
    processDetails.className = 'stats-details';
    processDetails.innerHTML = `
        <div class="stats-section">
            <div class="stats-header">
                <h3>待处理内容</h3>
                <span class="stats-badge highlight">${totalItems}条记录</span>
            </div>
            <div class="stats-cards-row">
                <div class="stat-card highlight">
                    <div class="stat-card-header">
                        <div class="stat-card-title">待处理记录</div>
                        <div class="stat-card-icon">
                            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"></path>
                                <rect x="8" y="2" width="8" height="4" rx="1" ry="1"></rect>
                            </svg>
                        </div>
                    </div>
                    <div class="stat-card-value">${totalItems}</div>
                    <div class="stat-card-progress">
                        <div class="progress-bar" style="width: 100%"></div>
                    </div>
                </div>

                <div class="stat-card update">
                    <div class="stat-card-header">
                        <div class="stat-card-title">数据更新</div>
                        <div class="stat-card-icon">
                            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <path d="M21 2v6h-6"></path>
                                <path d="M3 12a9 9 0 0 1 15-6.7L21 8"></path>
                                <path d="M3 22v-6h6"></path>
                                <path d="M21 12a9 9 0 0 1-15 6.7L3 16"></path>
                            </svg>
                        </div>
                    </div>
                    <div class="stat-card-value">${differences.length}</div>
                    <div class="stat-card-progress">
                        <div class="progress-bar" style="width: ${totalItems > 0 ? (differences.length / totalItems) * 100 : 0}%"></div>
                    </div>
                </div>

                <div class="stat-card new">
                    <div class="stat-card-header">
                        <div class="stat-card-title">新增学生</div>
                        <div class="stat-card-icon">
                            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <path d="M16 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path>
                                <circle cx="8.5" cy="7" r="4"></circle>
                                <line x1="20" y1="8" x2="20" y2="14"></line>
                                <line x1="23" y1="11" x2="17" y2="11"></line>
                            </svg>
                        </div>
                    </div>
                    <div class="stat-card-value">${newStudents.length}</div>
                    <div class="stat-card-progress">
                        <div class="progress-bar" style="width: ${totalItems > 0 ? (newStudents.length / totalItems) * 100 : 0}%"></div>
                    </div>
                </div>
            </div>
        </div>
    `;
    statsContainer.appendChild(processDetails);
    
    // 将统计卡片添加到容器
    statsCards.appendChild(statsContainer);
    
    // 创建分段控制器/标签页
    const resultTabs = document.getElementById('result-tabs');
    resultTabs.innerHTML = '';
    
    // 全部标签
    const allTab = document.createElement('div');
    allTab.className = 'segment-tab active';
    allTab.dataset.filter = 'all';
    allTab.innerHTML = `
        全部 <span class="segment-tab-count">${totalItems}</span>
    `;
    resultTabs.appendChild(allTab);
    
    // 更新标签
    const updateTab = document.createElement('div');
    updateTab.className = 'segment-tab';
    updateTab.dataset.filter = 'update';
    updateTab.innerHTML = `
        更新 <span class="segment-tab-count">${differences.length}</span>
    `;
    resultTabs.appendChild(updateTab);
    
    // 新增标签
    const newTab = document.createElement('div');
    newTab.className = 'segment-tab';
    newTab.dataset.filter = 'new';
    newTab.innerHTML = `
        新增 <span class="segment-tab-count">${newStudents.length}</span>
    `;
    resultTabs.appendChild(newTab);
    
    // 为标签添加点击事件
    const tabs = resultTabs.querySelectorAll('.segment-tab');
    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            // 移除所有标签的active类
            tabs.forEach(t => t.classList.remove('active'));
            // 为当前标签添加active类
            tab.classList.add('active');
            
            // 根据过滤条件显示/隐藏表格行
            const filter = tab.dataset.filter;
            const rows = differencesTableContainer.querySelectorAll('tr[data-type]');
            
            // 首先隐藏所有行
            rows.forEach(row => {
                row.style.display = 'none';
            });
            
            // 然后根据筛选条件显示行
            rows.forEach(row => {
                if (filter === 'all' || row.dataset.type === filter) {
                    row.style.display = '';
                }
            });
        });
    });
    
    // 创建表格
    const table = document.createElement('table');
    table.className = 'differences-table';
    
    // 创建表头
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    
    ['选择', '记录 ID', '姓名', '班级', '变更内容', '旧值', '新值', '类型'].forEach(text => {
        const th = document.createElement('th');
        th.textContent = text;
        headerRow.appendChild(th);
    });
    
    thead.appendChild(headerRow);
    table.appendChild(thead);
    
    // 创建表格主体
    const tbody = document.createElement('tbody');
    
    // 添加更新记录 - 多行展示方式
    differences.forEach((diff, index) => {
        let isFirstRow = true;
        
        // 为每个差异字段创建一行
        diff.differences.forEach(fieldDiff => {
            const row = document.createElement('tr');
            row.dataset.type = 'update';
            
            // 检查是否是新字段（A表中不存在，B表中存在的字段）
            const isNewField = fieldDiff.oldValue === undefined && fieldDiff.newValue !== undefined;
            
            // 如果是这个学生的第一行记录
            if (isFirstRow) {
                // 复选框单元格
                const checkboxCell = document.createElement('td');
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.checked = true;
                checkbox.dataset.diffIndex = index;
                checkbox.dataset.diffType = 'update';
                checkbox.className = 'diff-checkbox';
                checkboxCell.appendChild(checkbox);
                row.appendChild(checkboxCell);
                
                // ID单元格
                const idCell = document.createElement('td');
                idCell.textContent = diff.record || '';
                row.appendChild(idCell);
                
                // 姓名单元格
                const nameCell = document.createElement('td');
                nameCell.textContent = diff.name || '';
                row.appendChild(nameCell);
                
                // 班级单元格
                const classCell = document.createElement('td');
                classCell.textContent = diff.class || '';
                row.appendChild(classCell);
                
                // 字段名称
                const fieldCell = document.createElement('td');
                if (isNewField) {
                    fieldCell.className = 'new-field'; // 新字段样式
                }
                fieldCell.textContent = fieldDiff.field || '';
                row.appendChild(fieldCell);
                
                // 旧值单元格
                const oldValueCell = document.createElement('td');
                const oldValueSpan = document.createElement('span');
                oldValueSpan.className = isNewField ? 'empty-value' : 'old-value';
                oldValueSpan.textContent = fieldDiff.oldValue !== undefined ? fieldDiff.oldValue : '';
                oldValueCell.appendChild(oldValueSpan);
                row.appendChild(oldValueCell);
                
                // 新值单元格
                const newValueCell = document.createElement('td');
                const newValueSpan = document.createElement('span');
                newValueSpan.className = isNewField ? 'new-field-value' : 'new-value';
                newValueSpan.textContent = fieldDiff.newValue !== undefined ? fieldDiff.newValue : '';
                newValueCell.appendChild(newValueSpan);
                row.appendChild(newValueCell);
                
                // 类型单元格
                const typeCell = document.createElement('td');
                const updateTag = document.createElement('span');
                updateTag.className = isNewField ? 'tag tag-new-field' : 'tag tag-update';
                updateTag.textContent = isNewField ? '新数据' : '更新';
                typeCell.appendChild(updateTag);
                row.appendChild(typeCell);
                
                isFirstRow = false;
            } else {
                // 非第一行，前面单元格留空
                for (let i = 0; i < 4; i++) {
                    const emptyCell = document.createElement('td');
                    row.appendChild(emptyCell);
                }
                
                // 字段名称
                const fieldCell = document.createElement('td');
                if (isNewField) {
                    fieldCell.className = 'new-field'; // 新字段样式
                }
                fieldCell.textContent = fieldDiff.field || '';
                row.appendChild(fieldCell);
                
                // 旧值单元格
                const oldValueCell = document.createElement('td');
                const oldValueSpan = document.createElement('span');
                oldValueSpan.className = isNewField ? 'empty-value' : 'old-value';
                oldValueSpan.textContent = fieldDiff.oldValue !== undefined ? fieldDiff.oldValue : '';
                oldValueCell.appendChild(oldValueSpan);
                row.appendChild(oldValueCell);
                
                // 新值单元格
                const newValueCell = document.createElement('td');
                const newValueSpan = document.createElement('span');
                newValueSpan.className = isNewField ? 'new-field-value' : 'new-value';
                newValueSpan.textContent = fieldDiff.newValue !== undefined ? fieldDiff.newValue : '';
                newValueCell.appendChild(newValueSpan);
                row.appendChild(newValueCell);
                
                // 类型单元格（空）
                const typeCell = document.createElement('td');
                if (isNewField) {
                    const newFieldTag = document.createElement('span');
                    newFieldTag.className = 'tag tag-new-field';
                    newFieldTag.textContent = '新数据';
                    typeCell.appendChild(newFieldTag);
                }
                row.appendChild(typeCell);
            }
            
            tbody.appendChild(row);
        });
    });
    
    // 添加新学生记录
    newStudents.forEach((newStudent, index) => {
        const row = document.createElement('tr');
        row.dataset.type = 'new';
        row.className = 'added-row'; // 添加背景色
        
        // 复选框单元格
        const checkboxCell = document.createElement('td');
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.checked = true;
        checkbox.dataset.newStudentIndex = index;
        checkbox.dataset.diffType = 'new';
        checkbox.className = 'diff-checkbox';
        checkboxCell.appendChild(checkbox);
        row.appendChild(checkboxCell);
        
        // ID单元格
        const idCell = document.createElement('td');
        idCell.textContent = newStudent.record || '';
        row.appendChild(idCell);
        
        // 姓名单元格
        const nameCell = document.createElement('td');
        nameCell.textContent = newStudent.name || '';
        row.appendChild(nameCell);
        
        // 班级单元格
        const classCell = document.createElement('td');
        classCell.textContent = newStudent.class || '';
        row.appendChild(classCell);
        
        // 变更内容单元格
        const fieldCell = document.createElement('td');
        const addContent = document.createElement('div');
        addContent.className = 'flex items-center space-x-2';
        addContent.innerHTML = `
            <svg class="add-icon" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <circle cx="12" cy="12" r="10"></circle>
                <line x1="12" y1="8" x2="12" y2="16"></line>
                <line x1="8" y1="12" x2="16" y2="12"></line>
            </svg>
            <span class="add-text">整行新增</span>
        `;
        fieldCell.appendChild(addContent);
        row.appendChild(fieldCell);
        
        // 添加空的旧值和新值单元格
        const oldValueCell = document.createElement('td');
        row.appendChild(oldValueCell);
        
        const newValueCell = document.createElement('td');
        row.appendChild(newValueCell);
        
        // 类型单元格
        const typeCell = document.createElement('td');
        const newTag = document.createElement('span');
        newTag.className = 'tag tag-new';
        newTag.textContent = '新增';
        typeCell.appendChild(newTag);
        row.appendChild(typeCell);
        
        tbody.appendChild(row);
    });
    
    table.appendChild(tbody);
    differencesTableContainer.appendChild(table);
    
    // 确保确认按钮内容正确，添加图标
    const confirmBtn = document.getElementById('confirmUpdateBtn');
    if (confirmBtn) {
        confirmBtn.innerHTML = `
            确认并导出 (${totalItems})
        `;
    }
    
    // 显示比对结果容器
    differencesContainer.classList.remove('hidden');
}

/**
 * 滚动到更新按钮并高亮显示
 */
function scrollToUpdateButton() {
    // 滚动到确认更新按钮位置
    confirmUpdateBtn.scrollIntoView({ behavior: 'smooth', block: 'center' });
    
    // 添加高亮效果
    confirmUpdateBtn.classList.add('highlight-button');
    
    // 3秒后移除高亮效果
    setTimeout(() => {
        confirmUpdateBtn.classList.remove('highlight-button');
    }, 3000);
}

/**
 * 按工作表分组差异和新学生数据
 * @param {Array} differences - 差异数据数组
 * @param {Array} newStudents - 新学生数据数组
 * @returns {Object} - 按工作表分组的数据
 */
function groupDifferencesBySheet(differences, newStudents) {
    const result = {};
    
    // 分组差异数据
    differences.forEach(diff => {
        const sheetName = diff.rowA && diff.rowA['_原始工作表'] ? diff.rowA['_原始工作表'] : '未知工作表';
        
        if (!result[sheetName]) {
            result[sheetName] = {
                differences: [],
                newStudents: []
            };
        }
        
        result[sheetName].differences.push(diff);
    });
    
    // 分组新学生数据
    newStudents.forEach(student => {
        const sheetName = student.rowB && student.rowB['_原始工作表'] ? student.rowB['_原始工作表'] : '新增学生';
        
        if (!result[sheetName]) {
            result[sheetName] = {
                differences: [],
                newStudents: []
            };
        }
        
        result[sheetName].newStudents.push(student);
    });
    
    return result;
}

/**
 * 创建工作表内容视图
 * @param {string} sheetId - 工作表ID
 * @param {Array} differences - 差异数据数组
 * @param {Array} newStudents - 新学生数据数组
 * @returns {HTMLElement} - 内容视图元素
 */
function createSheetContentView(sheetId, differences, newStudents) {
    const container = document.createElement('div');
    container.className = 'sheet-content';
    container.dataset.sheetId = sheetId;
    
    // 明确设置data-sheet-id属性，以确保与选项卡的dataset属性匹配
    container.setAttribute('data-sheet-id', sheetId);
    
    const totalItems = differences.length + newStudents.length;
    
    if (totalItems === 0) {
        container.innerHTML = '<p>没有发现需要更新的数据</p>';
        return container;
    }
    
    // 创建筛选标签
    const filterTagsDiv = document.createElement('div');
    filterTagsDiv.className = 'filter-tags';
    
    const allTag = document.createElement('span');
    allTag.className = 'filter-tag active';
    allTag.textContent = `全部 (${totalItems})`;
    allTag.dataset.filter = 'all';
    
    const updateTag = document.createElement('span');
    updateTag.className = 'filter-tag';
    updateTag.textContent = `已更新 (${differences.length})`;
    updateTag.dataset.filter = 'update';
    
    const newTag = document.createElement('span');
    newTag.className = 'filter-tag new-student';
    newTag.textContent = `新增 (${newStudents.length})`;
    newTag.dataset.filter = 'new';
    
    filterTagsDiv.appendChild(allTag);
    filterTagsDiv.appendChild(updateTag);
    filterTagsDiv.appendChild(newTag);
    
    container.appendChild(filterTagsDiv);
    
    // 创建表格
    const table = document.createElement('table');
    table.className = 'differences-table';
    
    // 创建表头
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    
    ['选择', '记录 ID', '姓名', '班级', '变更内容', '旧值', '新值', '类型'].forEach(text => {
        const th = document.createElement('th');
        th.textContent = text;
        headerRow.appendChild(th);
    });
    
    thead.appendChild(headerRow);
    table.appendChild(thead);
    
    // 创建表格主体
    const tbody = document.createElement('tbody');
    
    // 添加差异记录
    differences.forEach((diff, index) => {
        // 为每个差异字段创建一行
        diff.differences.forEach((fieldDiff, fieldIndex) => {
            const row = document.createElement('tr');
            row.dataset.type = 'update'; // 标记为更新类型
            
            // 检查是否是新字段（A表中不存在，B表中存在的字段）
            const isNewField = fieldDiff.oldValue === undefined && fieldDiff.newValue !== undefined;
            
            // 如果是该记录的第一行，添加复选框、ID、姓名和班级
            if (fieldIndex === 0) {
                // 复选框单元格
                const checkboxCell = document.createElement('td');
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.checked = true;
                checkbox.dataset.diffIndex = index;
                checkbox.dataset.diffType = 'update';
                checkbox.className = 'diff-checkbox';
                checkboxCell.appendChild(checkbox);
                row.appendChild(checkboxCell);
                
                // ID单元格
                const idCell = document.createElement('td');
                idCell.textContent = diff.record;
                row.appendChild(idCell);
                
                // 姓名单元格
                const nameCell = document.createElement('td');
                nameCell.textContent = diff.name;
                row.appendChild(nameCell);
                
                // 班级单元格
                const classCell = document.createElement('td');
                classCell.textContent = diff.class;
                row.appendChild(classCell);
            } else {
                // 如果不是第一行，前面的单元格合并
                ['', '', '', ''].forEach(() => {
                    const emptyCell = document.createElement('td');
                    row.appendChild(emptyCell);
                });
            }
            
            // 字段名称单元格
            const fieldCell = document.createElement('td');
            if (isNewField) {
                fieldCell.className = 'new-field'; // 新字段样式
            }
            fieldCell.textContent = fieldDiff.field;
            row.appendChild(fieldCell);
            
            // 旧值单元格
            const oldValueCell = document.createElement('td');
            const oldValueSpan = document.createElement('span');
            oldValueSpan.className = isNewField ? 'empty-value' : 'old-value';
            oldValueSpan.textContent = fieldDiff.oldValue !== undefined ? fieldDiff.oldValue : '';
            oldValueCell.appendChild(oldValueSpan);
            row.appendChild(oldValueCell);
            
            // 新值单元格
            const newValueCell = document.createElement('td');
            const newValueSpan = document.createElement('span');
            newValueSpan.className = isNewField ? 'new-field-value' : 'new-value';
            newValueSpan.textContent = fieldDiff.newValue !== undefined ? fieldDiff.newValue : '';
            newValueCell.appendChild(newValueSpan);
            row.appendChild(newValueCell);
            
            // 类型单元格
            const typeCell = document.createElement('td');
            const updateTag = document.createElement('span');
            updateTag.className = isNewField ? 'tag tag-new-field' : 'tag tag-update';
            updateTag.textContent = isNewField ? '新数据' : '更新';
            typeCell.appendChild(updateTag);
            row.appendChild(typeCell);
            
            tbody.appendChild(row);
        });
    });
    
    // 添加新学生记录
    newStudents.forEach((newStudent, index) => {
        const row = document.createElement('tr');
        row.className = 'new-student-row'; // 添加新学生样式
        row.dataset.type = 'new'; // 标记为新增类型
        
        // 复选框单元格
        const checkboxCell = document.createElement('td');
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.checked = true;
        checkbox.dataset.newStudentIndex = index;
        checkbox.dataset.diffType = 'new';
        checkbox.className = 'diff-checkbox new-student-checkbox';
        checkbox.dataset.sheetId = sheetId;
        checkboxCell.appendChild(checkbox);
        row.appendChild(checkboxCell);
        
        // ID单元格
        const idCell = document.createElement('td');
        idCell.textContent = newStudent.record;
        row.appendChild(idCell);
        
        // 姓名单元格
        const nameCell = document.createElement('td');
        nameCell.textContent = newStudent.name;
        row.appendChild(nameCell);
        
        // 班级单元格
        const classCell = document.createElement('td');
        classCell.textContent = newStudent.class;
        row.appendChild(classCell);
        
        // 字段名称单元格 - 整行新增
        const fieldCell = document.createElement('td');
        fieldCell.textContent = '整行新增';
        row.appendChild(fieldCell);
        
        // 添加空的旧值和新值单元格
        const oldValueCell = document.createElement('td');
        row.appendChild(oldValueCell);
        
        const newValueCell = document.createElement('td');
        row.appendChild(newValueCell);
        
        // 类型单元格
        const typeCell = document.createElement('td');
        const newTag = document.createElement('span');
        newTag.className = 'tag tag-new';
        newTag.textContent = '新增';
        typeCell.appendChild(newTag);
        row.appendChild(typeCell);
        
        tbody.appendChild(row);
    });
    
    table.appendChild(tbody);
    
    // 添加全选/取消全选按钮
    const controlsDiv = document.createElement('div');
    controlsDiv.className = 'diff-controls';
    
    const selectAllBtn = document.createElement('button');
    selectAllBtn.type = 'button';
    selectAllBtn.textContent = '全选';
    selectAllBtn.className = 'select-all-btn';
    selectAllBtn.onclick = () => {
        container.querySelectorAll('.diff-checkbox').forEach(cb => cb.checked = true);
    };
    
    const deselectAllBtn = document.createElement('button');
    deselectAllBtn.type = 'button';
    deselectAllBtn.textContent = '取消全选';
    deselectAllBtn.className = 'select-all-btn';
    deselectAllBtn.style.marginLeft = '10px';
    deselectAllBtn.onclick = () => {
        container.querySelectorAll('.diff-checkbox').forEach(cb => cb.checked = false);
    };
    
    controlsDiv.appendChild(selectAllBtn);
    controlsDiv.appendChild(deselectAllBtn);
    
    container.appendChild(controlsDiv);
    container.appendChild(table);
    
    // 绑定筛选标签点击事件
    filterTagsDiv.querySelectorAll('.filter-tag').forEach(tag => {
        tag.addEventListener('click', () => {
            // 激活当前标签
            filterTagsDiv.querySelectorAll('.filter-tag').forEach(t => t.classList.remove('active'));
            tag.classList.add('active');
            
            const filter = tag.dataset.filter;
            
            // 根据筛选条件显示/隐藏行
            if (filter === 'all') {
                table.querySelectorAll('tbody tr').forEach(row => {
                    row.style.display = '';
                });
            } else {
                table.querySelectorAll('tbody tr').forEach(row => {
                    if (row.dataset.type === filter) {
                        row.style.display = '';
                    } else {
                        row.style.display = 'none';
                    }
                });
            }
        });
    });
    
    return container;
}

/**
 * 处理确认更新按钮点击事件
 */
function handleConfirmUpdate() {
    console.log('确认更新按钮被点击');
    
    if ((!globalComparisonResult.differences && !globalComparisonResult.newStudents) || !globalComparisonResult.originalDataA) {
        console.error('错误：没有可用的比对结果', {
            differences: globalComparisonResult.differences,
            newStudents: globalComparisonResult.newStudents,
            originalDataA: globalComparisonResult.originalDataA
        });
        statusDiv.textContent = '错误：没有可用的比对结果';
        statusDiv.style.color = 'red';
        return;
    }
    
    console.log('比对结果验证通过，开始获取选中项');
    
    // 获取所有选中的差异索引和新学生索引
    const selectedItems = {
        updateItems: [],
        newStudentItems: []
    };
    
    // 获取当前激活的工作表内容视图
    const activeContent = document.querySelector('.sheet-content.active');
    console.log('当前激活的工作表视图:', activeContent);
    
    if (!activeContent) {
        console.error('错误：无法确定当前激活的工作表视图');
        statusDiv.textContent = '错误：无法确定当前激活的工作表视图';
        statusDiv.style.color = 'red';
        return;
    }
    
    // 获取当前视图的工作表ID
    const activeSheetId = activeContent.dataset.sheetId;
    console.log('当前激活的工作表ID:', activeSheetId);
    
    // 从当前视图或所有视图获取选中的差异和新学生
    const getCheckedItems = (selector) => {
        // 如果是"全部"视图，合并所有视图中的选中项
        if (activeSheetId === 'all') {
            return document.querySelectorAll(selector);
        } else {
            // 否则只获取当前视图中的选中项
            return activeContent.querySelectorAll(selector);
        }
    };
    
    // 获取选中的更新项
    const checkedUpdateItems = getCheckedItems('.diff-checkbox[data-diff-type="update"]:checked');
    console.log(`找到 ${checkedUpdateItems.length} 个选中的更新项`);
    
    checkedUpdateItems.forEach(checkbox => {
        const diffIndex = parseInt(checkbox.dataset.diffIndex);
        console.log(`处理更新项 diffIndex=${diffIndex}`);
        
        // 检查是否为当前工作表或全部工作表视图
        if (!isNaN(diffIndex)) {
            if (activeSheetId === 'all') {
                // 全部视图 - 使用全局差异数组
                if (globalComparisonResult.differences[diffIndex]) {
                    selectedItems.updateItems.push(globalComparisonResult.differences[diffIndex]);
                    console.log(`添加全局差异项 #${diffIndex} 到更新列表`);
                } else {
                    console.warn(`警告：未找到全局差异项 #${diffIndex}`);
                }
            } else {
                // 单个工作表视图 - 使用特定工作表的差异数组
                if (globalComparisonResult.bySheet && 
                    globalComparisonResult.bySheet[activeSheetId] && 
                    globalComparisonResult.bySheet[activeSheetId].differences[diffIndex]) {
                    selectedItems.updateItems.push(globalComparisonResult.bySheet[activeSheetId].differences[diffIndex]);
                    console.log(`添加工作表 ${activeSheetId} 的差异项 #${diffIndex} 到更新列表`);
                } else {
                    console.warn(`警告：未找到工作表 ${activeSheetId} 的差异项 #${diffIndex}`, {
                        bySheet: !!globalComparisonResult.bySheet,
                        sheetExists: globalComparisonResult.bySheet ? !!globalComparisonResult.bySheet[activeSheetId] : false,
                        differencesExists: globalComparisonResult.bySheet && globalComparisonResult.bySheet[activeSheetId] ? 
                            !!globalComparisonResult.bySheet[activeSheetId].differences : false
                    });
                }
            }
        } else {
            console.warn(`警告：无效的差异索引 ${checkbox.dataset.diffIndex}`);
        }
    });
    
    // 获取选中的新学生
    const checkedNewStudents = getCheckedItems('.diff-checkbox[data-diff-type="new"]:checked');
    console.log(`找到 ${checkedNewStudents.length} 个选中的新学生项`);
    
    checkedNewStudents.forEach(checkbox => {
        const newStudentIndex = parseInt(checkbox.dataset.newStudentIndex);
        console.log(`处理新学生项 newStudentIndex=${newStudentIndex}`);
        
        // 检查是否为当前工作表或全部工作表视图
        if (!isNaN(newStudentIndex)) {
            if (activeSheetId === 'all') {
                // 全部视图 - 使用全局新学生数组
                if (globalComparisonResult.newStudents[newStudentIndex]) {
                    selectedItems.newStudentItems.push(globalComparisonResult.newStudents[newStudentIndex]);
                    console.log(`添加全局新学生项 #${newStudentIndex} 到更新列表`);
                } else {
                    console.warn(`警告：未找到全局新学生项 #${newStudentIndex}`);
                }
            } else {
                // 单个工作表视图 - 使用特定工作表的新学生数组
                if (globalComparisonResult.bySheet && 
                    globalComparisonResult.bySheet[activeSheetId] && 
                    globalComparisonResult.bySheet[activeSheetId].newStudents[newStudentIndex]) {
                    selectedItems.newStudentItems.push(globalComparisonResult.bySheet[activeSheetId].newStudents[newStudentIndex]);
                    console.log(`添加工作表 ${activeSheetId} 的新学生项 #${newStudentIndex} 到更新列表`);
                } else {
                    console.warn(`警告：未找到工作表 ${activeSheetId} 的新学生项 #${newStudentIndex}`, {
                        bySheet: !!globalComparisonResult.bySheet,
                        sheetExists: globalComparisonResult.bySheet ? !!globalComparisonResult.bySheet[activeSheetId] : false,
                        newStudentsExists: globalComparisonResult.bySheet && globalComparisonResult.bySheet[activeSheetId] ? 
                            !!globalComparisonResult.bySheet[activeSheetId].newStudents : false
                    });
                }
            }
        } else {
            console.warn(`警告：无效的新学生索引 ${checkbox.dataset.newStudentIndex}`);
        }
    });
    
    console.log('选中项汇总:', selectedItems);
    
    // 如果没有选中任何记录，提示用户
    if (selectedItems.updateItems.length === 0 && selectedItems.newStudentItems.length === 0) {
        console.warn('未选中任何记录');
        statusDiv.textContent = '请至少选择一条需要更新或添加的记录';
        statusDiv.style.color = 'orange';
        return;
    }
    
    // 创建进度指示器
    showProgressBar(true);
    console.log('显示进度条');
    
    // 创建进度回调函数
    const updateProgress = (message, percent) => {
        console.log(`进度更新: ${message} (${percent}%)`);
        const progressBar = document.querySelector('.progress-bar');
        if (progressBar) {
            statusDiv.textContent = message;
            progressBar.style.width = `${percent}%`;
        }
    };
    
    updateProgress('开始更新数据...', 10);
    
    // 备份当前状态以支持撤销
    globalComparisonResult.lastUpdateState = {
        originalDataA: structuredClone(globalComparisonResult.originalDataA),
        updatedDataA: globalComparisonResult.updatedDataA ? structuredClone(globalComparisonResult.updatedDataA) : null,
        differences: structuredClone(globalComparisonResult.differences || []),
        newStudents: structuredClone(globalComparisonResult.newStudents || [])
    };
    console.log('已备份当前状态以支持撤销功能');
    
    // 更新数据 - 使用Promise方式
    console.log('开始调用updateRecords函数更新数据');
    updateRecords(
        globalComparisonResult.originalDataA, 
        {
            differences: selectedItems.updateItems,
            newStudents: selectedItems.newStudentItems
        },
        updateProgress
    ).then(updatedDataA => {
        console.log('数据更新成功，获得更新后的数据:', {
            recordCount: updatedDataA.length
        });
        
        // 更新全局结果
        globalComparisonResult.updatedDataA = updatedDataA;
        globalComparisonResult.readyToExport = true;
        
        // 启用撤销按钮
        if (undoUpdateBtn) {
            undoUpdateBtn.disabled = false;
            console.log('已启用撤销按钮');
        }
        
        // 构建状态消息
        let updateMessage = '';
        if (selectedItems.updateItems.length > 0 && selectedItems.newStudentItems.length > 0) {
            updateMessage = `已确认更新 ${selectedItems.updateItems.length} 条记录并添加 ${selectedItems.newStudentItems.length} 名新学生`;
        } else if (selectedItems.updateItems.length > 0) {
            updateMessage = `已确认更新 ${selectedItems.updateItems.length} 条记录`;
        } else {
            updateMessage = `已确认添加 ${selectedItems.newStudentItems.length} 名新学生`;
        }
        
        // 显示状态
        updateProgress(`${updateMessage}，正在生成Excel文件...`, 98);
        
        // 导出Excel
        console.log('开始调用exportToExcel导出Excel文件');
        exportToExcel(updatedDataA, NAME_COLUMN);
        
        // 更新结果摘要
        let summaryHTML = `<p>数据更新完成！已生成更新后的Excel文件。</p>`;
        
        if (selectedItems.updateItems.length > 0) {
            summaryHTML += `<p>已更新 ${selectedItems.updateItems.length} 条记录</p>`;
        }
        
        if (selectedItems.newStudentItems.length > 0) {
            summaryHTML += `<p>已添加 ${selectedItems.newStudentItems.length} 名新学生</p>`;
        }
        
        summaryHTML += `<p>可以点击"撤销上一次更新"按钮撤销此次更新。</p>`;
        
        summaryContent.innerHTML = summaryHTML;
        console.log('已更新结果摘要');
        
        // 完成
        updateProgress(`更新完成！${updateMessage}，Excel文件已开始下载`, 100);
        statusDiv.style.color = 'green';
        
        // 隐藏进度条
        setTimeout(() => {
            showProgressBar(false);
            console.log('已隐藏进度条');
        }, 1500);
    }).catch(error => {
        console.error('更新过程中出错:', error);
        updateProgress('更新过程中出错: ' + error.message, 0);
        statusDiv.style.color = 'red';
        setTimeout(() => {
            showProgressBar(false);
        }, 3000);
    });
}

/**
 * 处理撤销更新按钮点击事件
 */
function handleUndoUpdate() {
    // 显示进度条
    showProgressBar(true);
    updateProgress('准备撤销上一次更新...', 10);
    
    // 撤销上一次更新
    undoLastUpdate(updateProgress)
        .then(success => {
            if (success) {
                updateProgress('撤销操作完成', 100);
                // 更新结果显示
                summaryContent.innerHTML = `<p>已成功撤销上一次更新操作。</p>`;
                
                // 禁用撤销按钮
                undoUpdateBtn.disabled = true;
            } else {
                updateProgress('没有可撤销的操作', 0);
                summaryContent.innerHTML = `<p>没有可撤销的操作。</p>`;
            }
            
            setTimeout(() => {
                showProgressBar(false);
            }, 1000);
        })
        .catch(error => {
            console.error('撤销过程中出错:', error);
            updateProgress('撤销过程中出错: ' + error.message, 0);
            setTimeout(() => {
                showProgressBar(false);
            }, 3000);
        });
}

/**
 * 撤销上一次更新操作
 */
function undoLastUpdate(progressCallback) {
    return new Promise(async (resolve) => {
        if (!globalComparisonResult.lastUpdateState) {
            if (progressCallback) progressCallback('没有可撤销的操作', 0);
            resolve(false);
            return;
        }
        
        if (progressCallback) progressCallback('正在撤销上一次更新...', 30);
        
        // 恢复上一次状态
        globalComparisonResult.originalDataA = globalComparisonResult.lastUpdateState.originalDataA;
        globalComparisonResult.updatedDataA = globalComparisonResult.lastUpdateState.updatedDataA;
        globalComparisonResult.differences = globalComparisonResult.lastUpdateState.differences;
        
        // 清除撤销状态
        globalComparisonResult.lastUpdateState = null;
        
        if (progressCallback) progressCallback('撤销完成', 100);
        
        // 重新显示差异
        displayDifferences(globalComparisonResult.differences);
        
        resolve(true);
    });
}

/**
 * 显示或隐藏进度条
 */
function showProgressBar(show) {
    // 清除之前的进度条
    statusDiv.innerHTML = '';
    
    if (show) {
        // 创建进度指示器
        const progressContainer = document.createElement('div');
        progressContainer.className = 'progress-container';
        const progressBar = document.createElement('div');
        progressBar.className = 'progress-bar';
        progressContainer.appendChild(progressBar);
        statusDiv.appendChild(progressContainer);
        
        progressBar.style.width = '0%';
        
        // 隐藏快速更新按钮
        quickUpdateBtn.classList.add('hidden');
    } else {
        // 确保进度条完全隐藏
        const progressBar = document.querySelector('.progress-bar');
        if (progressBar) {
            progressBar.style.width = '0%';
        }
    }
}

/**
 * 比较两组数据，找出所有不同之处
 * @param {Array<Object>} dataA - 表格 A 的数据
 * @param {Array<Object>} dataB - 表格 B 的数据
 * @param {Function} progressCallback - 进度回调函数
 * @returns {Array<Object>} - 带有状态标记的结果数据
 */
function compareData(dataA, dataB, progressCallback) {
    // 存储原始数据
    globalComparisonResult.originalDataA = [...dataA];
    globalComparisonResult.originalDataB = [...dataB];
    globalComparisonResult.keyColumn = NAME_COLUMN; // 使用姓名作为匹配依据
    
    // 查找所有需要更新的数据
    const differences = findDifferences(dataA, dataB, progressCallback);
    globalComparisonResult.differences = differences;
    
    console.log(`发现 ${differences.length} 条需要更新的记录`);
    
    // 显示差异表格
    if (differences.length > 0) {
        differencesContainer.classList.remove('hidden');
        displayDifferences(differences);
    } else {
        differencesContainer.classList.add('hidden');
    }
    
    // 记录原始数据长度，用于日志调试
    console.log(`表格A原始数据: ${dataA.length}条, 表格B原始数据: ${dataB.length}条`);
    
    // 检查数据中是否存在重复的主键值
    const keyValuesA = dataA.map(row => row[keyColumn] !== null && row[keyColumn] !== undefined ? String(row[keyColumn]).trim() : '');
    const keyValuesB = dataB.map(row => row[keyColumn] !== null && row[keyColumn] !== undefined ? String(row[keyColumn]).trim() : '');
    
    // 检查重复值
    const duplicatesA = findDuplicates(keyValuesA);
    const duplicatesB = findDuplicates(keyValuesB);
    
    if (duplicatesA.length > 0) {
        console.warn(`警告: 表格A中存在重复的主键值: ${duplicatesA.join(', ')}`);
    }
    
    if (duplicatesB.length > 0) {
        console.warn(`警告: 表格B中存在重复的主键值: ${duplicatesB.join(', ')}`);
    }
    
    // 规范化主键并将数据转换为Map，处理可能的格式差异
    const mapA = new Map();
    const mapB = new Map();
    
    // 处理表格A的数据，使用索引+值作为复合键以处理重复值
    dataA.forEach((row, index) => {
        // 确保主键值被标准化处理（转为字符串并去除空格）
        const keyValue = row[keyColumn] !== null && row[keyColumn] !== undefined 
            ? String(row[keyColumn]).trim() 
            : '';
        // 使用组合键以确保唯一性
        const uniqueKey = `${keyValue}_${index}`;
        mapA.set(uniqueKey, { row, originalKey: keyValue });
    });
    
    // 处理表格B的数据
    dataB.forEach((row, index) => {
        // 确保主键值被标准化处理（转为字符串并去除空格）
        const keyValue = row[keyColumn] !== null && row[keyColumn] !== undefined 
            ? String(row[keyColumn]).trim() 
            : '';
        // 使用组合键以确保唯一性
        const uniqueKey = `${keyValue}_${index}`;
        mapB.set(uniqueKey, { row, originalKey: keyValue });
    });
    
    // 创建反向查找映射，用于找到匹配的记录
    const keyToIndexMapA = new Map();
    const keyToIndexMapB = new Map();
    
    // 填充反向查找映射
    [...mapA.entries()].forEach(([uniqueKey, { originalKey }]) => {
        if (!keyToIndexMapA.has(originalKey)) {
            keyToIndexMapA.set(originalKey, []);
        }
        keyToIndexMapA.get(originalKey).push(uniqueKey);
    });
    
    [...mapB.entries()].forEach(([uniqueKey, { originalKey }]) => {
        if (!keyToIndexMapB.has(originalKey)) {
            keyToIndexMapB.set(originalKey, []);
        }
        keyToIndexMapB.get(originalKey).push(uniqueKey);
    });
    
    const result = [];
    
    // 首先处理表格A中的记录
    // 跟踪已处理过的B记录，避免重复匹配
    const processedKeysB = new Set();
    
    for (const [uniqueKeyA, { row: rowA, originalKey: keyA }] of mapA.entries()) {
        // 检查表格B中是否有对应的主键
        if (!keyToIndexMapB.has(keyA)) {
            // 表格B中没有这个主键，A独有记录
            result.push({
                ...rowA,
                '比对状态': '仅在表格A中存在',
                '差异说明': '该记录仅在表格A中存在'
            });
        } else {
            // B中有相同主键的记录，需要比较内容
            // 这里我们取B中该主键的第一个未处理的记录进行比较
            const matchingKeysB = keyToIndexMapB.get(keyA);
            let matchingKeyB = null;
            
            // 找出第一个未处理的B记录
            for (const key of matchingKeysB) {
                if (!processedKeysB.has(key)) {
                    matchingKeyB = key;
                    processedKeysB.add(key); // 标记为已处理
                    break;
                }
            }
            
            // 如果所有B记录都已处理，则使用第一个（这种情况表明A中有重复记录）
            if (!matchingKeyB) {
                matchingKeyB = matchingKeysB[0];
                console.log(`警告: A中的记录(${keyA})重复匹配B中已处理的记录`);
            }
            
            const { row: rowB } = mapB.get(matchingKeyB);
            
            // 获取两行的所有字段名
            const allFields = new Set([
                ...Object.keys(rowA),
                ...Object.keys(rowB)
            ]);
            
            let hasDifference = false;
            const differences = [];
            
            // 比较每个字段
            for (const field of allFields) {
                // 跳过比对状态和差异说明字段
                if (field === '比对状态' || field === '差异说明') continue;
                // 跳过主键列，因为我们已经知道它们匹配
                if (field === keyColumn) continue;
                
                // 如果选择了忽略班级差异，且字段名包含"班级"或为"班号"，则跳过
                if (ignoreClasses.checked && (field.includes('班级') || field === '班号')) {
                    continue;
                }
                
                const valueA = rowA[field];
                const valueB = rowB[field];
                
                // 规范化比较值（转为字符串并去除两端空格）
                const normalizedA = valueA !== null && valueA !== undefined ? String(valueA).trim() : '';
                const normalizedB = valueB !== null && valueB !== undefined ? String(valueB).trim() : '';
                
                // 如果规范化后的值不相等，记录差异
                if (normalizedA !== normalizedB) {
                    hasDifference = true;
                    differences.push(`${field}: A="${valueA}" B="${valueB}"`);
                }
            }
            
            // 如果有差异，记录详细信息
            if (hasDifference) {
                result.push({
                    ...rowA,
                    '比对状态': '数据存在差异',
                    '差异说明': differences.join('; ')
                });
            } else if (showIdentical.checked) {
                // 如果选中了"显示完全相同的记录"选项，则添加相同记录
                result.push({
                    ...rowA,
                    '比对状态': '完全相同',
                    '差异说明': '两表中数据完全一致'
                });
            }
            // 否则完全相同的记录不添加到结果中
        }
    }
    
    // 处理表格B中独有的记录或未处理的记录
    for (const [uniqueKeyB, { row: rowB, originalKey: keyB }] of mapB.entries()) {
        // 检查表格A中是否有对应的主键
        if (!keyToIndexMapA.has(keyB)) {
            // 表格A中没有这个主键，B独有记录
            result.push({
                ...rowB,
                '比对状态': '仅在表格B中存在',
                '差异说明': '该记录仅在表格B中存在'
            });
        } else if (!processedKeysB.has(uniqueKeyB)) {
            // 这条记录有对应的主键，但未被A中的记录匹配过
            // 这表明B中有多条相同主键的记录，但A中的相同主键记录较少
            result.push({
                ...rowB,
                '比对状态': '仅在表格B中存在',
                '差异说明': '表格B中存在额外的重复记录'
            });
        }
        // 如果已被处理过，则跳过
    }
    
    // 输出处理前的原始记录数量，帮助调试
    console.log(`处理前的记录数量: A表=${dataA.length}, B表=${dataB.length}`);

    // 排序：先显示已更新记录，然后是有差异的记录
    result.sort((a, b) => {
        const order = { 
            '已更新': 0,
            '数据存在差异': 1, 
            '仅在表格A中存在': 2, 
            '仅在表格B中存在': 3,
            '完全相同': 4
        };
        
        // 首先按班级分组排序（如果选择了按班级分组）
        if (groupByClass.checked) {
            // 尝试找到班级字段
            const classFields = ['班级', '班号', '班', 'class'];
            for (const field of classFields) {
                if (a[field] && b[field]) {
                    const classCompare = String(a[field]).localeCompare(String(b[field]));
                    if (classCompare !== 0) return classCompare;
                    break;
                }
            }
        }
        
        // 然后按状态排序
        const statusOrder = order[a['比对状态']] - order[b['比对状态']];
        if (statusOrder !== 0) return statusOrder;
        
        // 如果状态相同，尝试按考试号排序
        if (a['考试号'] && b['考试号']) {
            return String(a['考试号']).localeCompare(String(b['考试号']));
        }
        
        // 如果没有考试号，尝试按学号排序
        if (a['学号'] && b['学号']) {
            return String(a['学号']).localeCompare(String(b['学号']));
        }
        
        // 默认返回0（保持原顺序）
        return 0;
    });
    
    // 输出一些统计信息到控制台，帮助调试
    console.log(`比对完成：找到 ${result.length} 条不一致记录`);
    console.log(`- 数据存在差异: ${result.filter(r => r['比对状态'] === '数据存在差异').length} 条`);
    console.log(`- 仅在表格A中存在: ${result.filter(r => r['比对状态'] === '仅在表格A中存在').length} 条`);
    console.log(`- 仅在表格B中存在: ${result.filter(r => r['比对状态'] === '仅在表格B中存在').length} 条`);
    
    return result;
}

/**
 * 将 JSON 数据导出为 Excel 文件并触发下载，支持筛选功能
 * @param {Array<Object>} data - 要导出的数据
 * @param {string} keyColumn - 主键列名
 */
function exportToExcel(data, keyColumn) {
    try {
        console.log('开始导出Excel文件，数据量:', data.length);
        
    if (data.length === 0) {
        statusDiv.textContent = "没有需要导出的数据。";
            console.error('导出失败：没有数据可导出');
        return;
    }
    
    // 数据去重 - 防止完全相同的记录重复显示
    const deduplicatedData = deduplicateResults(data);
    console.log(`原始结果: ${data.length}条, 去重后: ${deduplicatedData.length}条`);
        
        // 检查XLSX对象是否可用
        if (typeof XLSX === 'undefined') {
            statusDiv.textContent = '导出失败：Excel处理库(XLSX)未加载';
            statusDiv.style.color = 'red';
            console.error('导出失败：XLSX对象未定义');
            return;
        }
    
    // 创建一个新的工作簿
        console.log('创建工作簿...');
    const workbook = XLSX.utils.book_new();
    
    // 如果选择了按班级分组，且有班级字段，则尝试分表导出
        if (groupByClass && groupByClass.checked) {
            console.log('按班级分组导出...');
        const groupedData = groupDataByClass(deduplicatedData);
        
        if (groupedData.groups.size > 0) {
            // 添加总表
                console.log('添加总表...');
            const worksheet = XLSX.utils.json_to_sheet(deduplicatedData);
            addAutoFilter(worksheet);
                applyFormatting(worksheet, deduplicatedData); // 应用格式化（包括新学生黄色背景）
            XLSX.utils.book_append_sheet(workbook, worksheet, '所有数据');
            
            // 添加各班级表
                console.log(`添加${groupedData.groups.size}个班级表...`);
            for (const [className, records] of groupedData.groups.entries()) {
                if (records.length > 0) {
                    const classSheet = XLSX.utils.json_to_sheet(records);
                    addAutoFilter(classSheet);
                        applyFormatting(classSheet, records); // 应用格式化
                    // 限制工作表名长度，Excel限制为31个字符
                    const sheetName = className.length > 29 ? className.substring(0, 29) : className;
                        console.log(`添加班级表: ${sheetName}, 包含${records.length}条记录`);
                    XLSX.utils.book_append_sheet(workbook, classSheet, sheetName);
                }
            }
        } else {
            // 无法分组，仅添加总表
                console.log('无法按班级分组，添加单一工作表...');
            const worksheet = XLSX.utils.json_to_sheet(deduplicatedData);
            addAutoFilter(worksheet);
                applyFormatting(worksheet, deduplicatedData); // 应用格式化
            XLSX.utils.book_append_sheet(workbook, worksheet, '比对结果');
        }
    } else {
        // 不分组，只有一个表
            console.log('添加单一工作表...');
        const worksheet = XLSX.utils.json_to_sheet(deduplicatedData);
        addAutoFilter(worksheet);
            applyFormatting(worksheet, deduplicatedData); // 应用格式化
        XLSX.utils.book_append_sheet(workbook, worksheet, '比对结果');
    }

    // 生成文件名，包含日期
    const today = new Date();
    const dateStr = `${today.getFullYear()}-${(today.getMonth()+1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}`;
    const fileName = `学校数据比对结果_${dateStr}.xlsx`;

        console.log(`准备写入文件: ${fileName}`);
        
        // 使用try-catch专门捕获writeFile可能的错误
        try {
    // 生成并下载 Excel 文件
    XLSX.writeFile(workbook, fileName);
            console.log('Excel文件生成成功!');
            statusDiv.textContent = `Excel文件已生成: ${fileName}`;
            statusDiv.style.color = 'green';
        } catch (writeError) {
            console.error('写入Excel文件时出错:', writeError);
            statusDiv.textContent = `导出Excel文件失败: ${writeError.message}`;
            statusDiv.style.color = 'red';
            
            // 尝试使用备用方法导出
            try {
                console.log('尝试使用备用方法导出...');
                const blob = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
                const url = URL.createObjectURL(new Blob([blob], { type: 'application/octet-stream' }));
                
                // 创建一个下载链接并触发点击
                const downloadLink = document.createElement('a');
                downloadLink.href = url;
                downloadLink.download = fileName;
                document.body.appendChild(downloadLink);
                downloadLink.click();
                document.body.removeChild(downloadLink);
                
                console.log('备用导出方法成功!');
                statusDiv.textContent = `Excel文件已生成: ${fileName}`;
                statusDiv.style.color = 'green';
            } catch (altError) {
                console.error('备用导出方法也失败:', altError);
                statusDiv.textContent = `所有导出方法均失败，请检查浏览器设置或联系开发者`;
                statusDiv.style.color = 'red';
            }
        }
    } catch (error) {
        console.error('导出Excel过程中出错:', error);
        statusDiv.textContent = `导出Excel过程中出错: ${error.message}`;
        statusDiv.style.color = 'red';
    }
}

/**
 * 为工作表应用格式化（包括新学生的黄色背景）
 * @param {Object} worksheet - 工作表对象
 * @param {Array<Object>} data - 数据数组
 */
function applyFormatting(worksheet, data) {
    // 获取工作表范围
    if (!worksheet['!ref']) return;
    
    // 创建单元格样式对象
    if (!worksheet['!cols']) worksheet['!cols'] = [];
    
    // 自动调整列宽
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for (let i = 0; i <= range.e.c; i++) {
        worksheet['!cols'][i] = { wch: 15 }; // 默认列宽
    }
    
    // 应用单元格格式
    if (!worksheet['!comments']) worksheet['!comments'] = {};
    
    // 查找"_已更新"列的索引
    let updatedColIndex = -1;
    // 遍历表头找到"_已更新"列
    for (let col = 0; col <= range.e.c; col++) {
        const cellRef = XLSX.utils.encode_cell({r: 0, c: col});
        if (worksheet[cellRef] && worksheet[cellRef].v === '_已更新') {
            updatedColIndex = col;
            break;
        }
    }
    
    // 为每条记录应用格式
    for (let i = 0; i < data.length; i++) {
        // Excel行索引从1开始，第1行是标题
        const row = i + 1;
        
        // 处理"_已更新"列的值
        if (updatedColIndex !== -1) {
            const updatedCellRef = XLSX.utils.encode_cell({r: row, c: updatedColIndex});
            
            // 确保单元格存在
            if (!worksheet[updatedCellRef]) {
                worksheet[updatedCellRef] = { v: "" };
            }
            
            // 如果是新增学生，显示"新增"
            if (data[i]['_已更新'] === '新增') {
                worksheet[updatedCellRef].v = '新增';
            }
        }
        
        // 检查是否为新学生
        if (data[i]['_是新学生'] === true) {
            // 为该行的所有单元格添加黄色背景
            for (let col = 0; col <= range.e.c; col++) {
                const cellRef = XLSX.utils.encode_cell({r: row, c: col});
                
                // 确保单元格存在
                if (!worksheet[cellRef]) {
                    worksheet[cellRef] = { v: "" };
                }
                
                // 添加单元格样式（黄色背景）
                if (!worksheet[cellRef].s) worksheet[cellRef].s = {};
                worksheet[cellRef].s.fill = {
                    patternType: "solid",
                    fgColor: { rgb: "FFFACD" } // 浅黄色
                };
            }
        }
    }
}

/**
 * 添加自动筛选功能到工作表
 * @param {Object} worksheet - 工作表对象
 */
function addAutoFilter(worksheet) {
    // 获取工作表中的数据范围
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    // 设置筛选范围（从A1到最后一列的最后一行）
    worksheet['!autofilter'] = { ref: XLSX.utils.encode_range(
        {r: 0, c: 0}, 
        {r: range.e.r, c: range.e.c}
    )};
}

/**
 * 按班级字段分组数据
 * @param {Array<Object>} data - 要分组的数据
 * @returns {Object} - 包含分组结果和使用的班级字段
 */
function groupDataByClass(data) {
    // 尝试找出班级字段名
    const possibleClassFields = ['班级', '班号', '班', 'class'];
    let classField = null;
    
    // 在数据的第一条记录中查找可能的班级字段
    if (data.length > 0) {
        const firstRecord = data[0];
        for (const field of possibleClassFields) {
            if (firstRecord[field] !== undefined) {
                classField = field;
                break;
            }
        }
    }
    
    // 分组结果
    const result = {
        classField: classField,
        groups: new Map()
    };
    
    // 如果找不到班级字段，返回空结果
    if (!classField) return result;
    
    // 按班级分组
    for (const record of data) {
        const classValue = String(record[classField] || '未分类');
        if (!result.groups.has(classValue)) {
            result.groups.set(classValue, []);
        }
        result.groups.get(classValue).push(record);
    }
    
    return result;
}

/**
 * 查找数组中的重复项
 * @param {Array} arr - 要检查的数组
 * @returns {Array} - 重复项的数组
 */
function findDuplicates(arr) {
    const seen = new Set();
    const duplicates = new Set();
    
    for (const item of arr) {
        if (seen.has(item)) {
            duplicates.add(item);
        } else {
            seen.add(item);
        }
    }
    
    return Array.from(duplicates);
}

/**
 * 对比对结果数据进行去重处理
 * @param {Array<Object>} results - 比对结果数据
 * @returns {Array<Object>} - 去重后的结果数据
 */
function deduplicateResults(results) {
    // 使用JSON字符串作为键来检测完全相同的对象
    const uniqueResults = new Map();
    
    // 处理结果中的每一条记录
    results.forEach(result => {
        // 创建记录的唯一标识符 - 使用考试号和学校作为复合键
        let uniqueKey = '';
        if (result['考试号']) {
            uniqueKey = `${result['考试号']}_${result['学校'] || ''}_${result['姓名'] || ''}`;
        } else if (result['学号']) {
            uniqueKey = `${result['学号']}_${result['学校'] || ''}_${result['姓名'] || ''}`;
        } else {
            // 如果没有考试号或学号，使用JSON字符串作为键
            uniqueKey = JSON.stringify(result);
        }
        
        // 只保留每个唯一标识符的第一条记录
        if (!uniqueResults.has(uniqueKey)) {
            uniqueResults.set(uniqueKey, result);
        }
    });
    
    // 返回去重后的结果数组
    return Array.from(uniqueResults.values());
}

/**
 * 生成比对结果的摘要信息
 * @param {Array<Object>} results - 比对结果
 * @param {number} totalA - A表总记录数
 * @param {number} totalB - B表总记录数
 */
function generateResultsSummary(results, totalA, totalB) {
    // 统计工作表信息
    const sheetsInfo = analyzeSheetDistribution(results);
    // 显示摘要容器
    summaryDiv.classList.remove('hidden');
    
    // 按比对状态分类统计
    const countByStatus = {
        '已更新': results.filter(r => r['比对状态'] === '已更新').length,
        '数据存在差异': results.filter(r => r['比对状态'] === '数据存在差异').length,
        '仅在表格A中存在': results.filter(r => r['比对状态'] === '仅在表格A中存在').length,
        '仅在表格B中存在': results.filter(r => r['比对状态'] === '仅在表格B中存在').length,
        '完全相同': results.filter(r => r['比对状态'] === '完全相同').length
    };
    
    // 如果选择了按班级分组，尝试按班级分组统计
    let classSummary = '';
    if (groupByClass.checked) {
        const classSummaryData = groupResultsByClass(results);
        if (classSummaryData) {
            classSummary = `<h4>班级差异统计：</h4>
                <ul>
                    ${classSummaryData.map(item => 
                        `<li>${item.class}: 差异${item.diff}条，A独有${item.onlyA}条，B独有${item.onlyB}条</li>`
                    ).join('')}
                </ul>`;
        }
    }
    
    // 生成工作表分布信息
    let sheetsSummary = '';
    if (sheetsInfo.totalSheets > 0) {
        sheetsSummary = `
            <h4>工作表分布：</h4>
            <ul>
                ${sheetsInfo.sheets.map(sheet => 
                    `<li>${sheet.name}: 共${sheet.total}条记录 (差异${sheet.diff}条, A独有${sheet.onlyA}条, B独有${sheet.onlyB}条${showIdentical.checked ? `, 相同${sheet.identical}条` : ''})</li>`
                ).join('')}
            </ul>
        `;
    }

    // 构建摘要HTML
    summaryContent.innerHTML = `
        <div>
            <p><strong>表格A总记录数：</strong> ${totalA}条（来自${sheetsInfoA.selectedSheets.size}个工作表）</p>
            <p><strong>表格B总记录数：</strong> ${totalB}条（来自${sheetsInfoB.selectedSheets.size}个工作表）</p>
            <p><strong>比对结果：</strong></p>
            <ul>
                ${globalComparisonResult.differences && globalComparisonResult.differences.length > 0 ? 
                    `<li class="highlight-diff"><strong>需要更新的记录: ${globalComparisonResult.differences.length}条</strong></li>` : ''}
                ${countByStatus['已更新'] > 0 ? `<li>已更新记录: ${countByStatus['已更新']}条</li>` : ''}
                <li>数据存在差异: ${countByStatus['数据存在差异']}条</li>
                <li>仅在表格A中存在: ${countByStatus['仅在表格A中存在']}条</li>
                <li>仅在表格B中存在: ${countByStatus['仅在表格B中存在']}条</li>
                ${showIdentical.checked ? `<li>完全相同: ${countByStatus['完全相同']}条</li>` : ''}
            </ul>
            ${sheetsSummary}
            ${classSummary}
        </div>
    `;
}

/**
 * 按班级分组统计结果
 * @param {Array<Object>} results - 比对结果
 * @returns {Array|null} - 按班级分组的统计数据，如果无法分组则返回null
 */
function groupResultsByClass(results) {
    // 尝试找出班级字段名
    const possibleClassFields = ['班级', '班号', '班', 'class'];
    let classField = null;
    
    // 在结果的第一条记录中查找可能的班级字段
    if (results.length > 0) {
        const firstRecord = results[0];
        for (const field of possibleClassFields) {
            if (firstRecord[field] !== undefined) {
                classField = field;
                break;
            }
        }
    }
    
    // 如果找不到班级字段，返回null
    if (!classField) return null;
    
    // 按班级分组
    const classSummary = new Map();
    
    for (const record of results) {
        const classValue = record[classField] || '未知班级';
        if (!classSummary.has(classValue)) {
            classSummary.set(classValue, {
                class: classValue,
                diff: 0,
                onlyA: 0,
                onlyB: 0
            });
        }
        
        const stats = classSummary.get(classValue);
        
        // 更新统计
        if (record['比对状态'] === '数据存在差异') {
            stats.diff++;
        } else if (record['比对状态'] === '仅在表格A中存在') {
            stats.onlyA++;
        } else if (record['比对状态'] === '仅在表格B中存在') {
            stats.onlyB++;
        }
    }
    
    return Array.from(classSummary.values());
}

/**
 * 分析结果中工作表的分布情况
 * @param {Array<Object>} results - 比对结果
 * @returns {Object} - 工作表分布信息
 */
function analyzeSheetDistribution(results) {
    // 统计信息
    const sheetStats = {
        totalSheets: 0,
        sheetsA: 0,
        sheetsB: 0,
        sheets: [],
        sheetMap: new Map()
    };
    
    // 统计每个工作表的记录
    const sheetsA = new Set();
    const sheetsB = new Set();
    
    // 遍历所有结果记录
    for (const record of results) {
        // 获取工作表名
        const sheetName = record['_原始工作表'] || '未知工作表';
        
        // 更新工作表统计
        if (!sheetStats.sheetMap.has(sheetName)) {
            sheetStats.sheetMap.set(sheetName, {
                name: sheetName,
                total: 0,
                diff: 0,
                onlyA: 0,
                onlyB: 0,
                identical: 0
            });
        }
        
        const sheetInfo = sheetStats.sheetMap.get(sheetName);
        sheetInfo.total++;
        
        // 根据比对状态更新各项统计
        if (record['比对状态'] === '数据存在差异') {
            sheetInfo.diff++;
        } else if (record['比对状态'] === '仅在表格A中存在') {
            sheetInfo.onlyA++;
            sheetsA.add(sheetName);
        } else if (record['比对状态'] === '仅在表格B中存在') {
            sheetInfo.onlyB++;
            sheetsB.add(sheetName);
        } else if (record['比对状态'] === '完全相同') {
            sheetInfo.identical++;
        }
    }
    
    // 更新总体统计
    sheetStats.sheets = Array.from(sheetStats.sheetMap.values());
    sheetStats.totalSheets = sheetStats.sheetMap.size;
    sheetStats.sheetsA = sheetsA.size;
    sheetStats.sheetsB = sheetsB.size;
    
    return sheetStats;
}

/**
 * 读取Excel文件并返回workbook对象
 * @param {File} file - Excel文件
 * @returns {Promise<Object>} - XLSX workbook对象
 */
function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                resolve(workbook);
            } catch (error) {
                reject(error);
            }
        };
        reader.onerror = (error) => reject(error);
        reader.readAsArrayBuffer(file);
    });
}

/**
 * 按工作表分别比对并处理数据
 * @param {Object} dataBySheetA - A表按工作表组织的数据
 * @param {Object} dataBySheetB - B表按工作表组织的数据
 * @param {Function} progressCallback - 进度回调函数
 * @returns {Promise<Object>} - 按工作表组织的比对结果
 */
async function compareDataBySheets(dataBySheetA, dataBySheetB, progressCallback) {
    return new Promise(async (resolve) => {
        // 记录所有工作表的比对结果
        const allResults = {
            bySheet: {},
            combined: {
                differences: [],
                newStudents: []
            },
            sheetStats: []
        };
        
        // 获取两表中共有的工作表
        const sheetsA = Object.keys(dataBySheetA);
        const sheetsB = Object.keys(dataBySheetB);
        const commonSheets = sheetsA.filter(sheet => sheetsB.includes(sheet));
        
        // 如果没有共有的工作表，返回空结果
        if (commonSheets.length === 0) {
            if (progressCallback) progressCallback('没有找到两表中共有的工作表', 100);
            resolve(allResults);
            return;
        }
        
        // 记录处理进度
        let processedSheets = 0;
        const totalSheets = commonSheets.length;
        
        // 对每个共有工作表分别进行比对
        for (const sheetName of commonSheets) {
            if (progressCallback) {
                progressCallback(`正在比对工作表 ${sheetName}...`, 
                    Math.floor(50 + (processedSheets / totalSheets) * 50));
            }
            
            const dataA = dataBySheetA[sheetName] || [];
            const dataB = dataBySheetB[sheetName] || [];
            
            // 对当前工作表进行比对
            const diffResult = await findDifferences(dataA, dataB);
            
            // 记录此工作表的结果
            allResults.bySheet[sheetName] = diffResult;
            
            // 合并到总结果中
            allResults.combined.differences.push(...diffResult.differences);
            allResults.combined.newStudents.push(...diffResult.newStudents);
            
            // 添加工作表统计信息
            allResults.sheetStats.push({
                name: sheetName,
                recordsA: dataA.length,
                recordsB: dataB.length,
                differences: diffResult.differences.length,
                newStudents: diffResult.newStudents.length
            });
            
            // 更新进度
            processedSheets++;
        }
        
        resolve(allResults);
    });
}

/**
 * 显示带有工作表选项卡的差异结果
 * @param {Object} allSheetResults - 所有工作表的比对结果
 */
function displayDifferencesWithSheetTabs(allSheetResults) {
    // 清空容器
    differencesTableContainer.innerHTML = '';
    
    // 提取所有差异和新学生
    const combinedDifferences = allSheetResults.combined.differences || [];
    const combinedNewStudents = allSheetResults.combined.newStudents || [];
    const totalItems = combinedDifferences.length + combinedNewStudents.length;
    
    if (totalItems === 0) {
        differencesTableContainer.innerHTML = '<p>没有发现需要更新的数据</p>';
        return;
    }
    
    // 创建比对信息摘要区域
    const summaryDiv = document.createElement('div');
    summaryDiv.className = 'comparison-summary';
    
    // 获取A表和B表的记录数
    const totalA = globalComparisonResult.originalDataA ? globalComparisonResult.originalDataA.length : 0;
    const totalB = globalComparisonResult.originalDataB ? globalComparisonResult.originalDataB.length : 0;
    
    // 计算仅在A表存在的记录数（通过名称比较）
    const namesInB = new Set();
    if (globalComparisonResult.originalDataB) {
        globalComparisonResult.originalDataB.forEach(row => {
            if (row[NAME_COLUMN]) {
                namesInB.add(String(row[NAME_COLUMN]).trim());
            }
        });
    }
    
    let onlyInA = 0;
    if (globalComparisonResult.originalDataA) {
        globalComparisonResult.originalDataA.forEach(row => {
            if (row[NAME_COLUMN] && !namesInB.has(String(row[NAME_COLUMN]).trim())) {
                onlyInA++;
            }
        });
    }
    
    summaryDiv.innerHTML = `
        <h4>数据比对结果概览</h4>
        <div class="stats">
            <div class="stat-item">
                <div class="stat-label">表格A总记录</div>
                <div class="stat-value">${totalA}</div>
            </div>
            <div class="stat-item">
                <div class="stat-label">表格B总记录</div>
                <div class="stat-value">${totalB}</div>
            </div>
            <div class="stat-item highlight">
                <div class="stat-label">需要更新记录</div>
                <div class="stat-value">${combinedDifferences.length}</div>
            </div>
            <div class="stat-item highlight">
                <div class="stat-label">新增学生</div>
                <div class="stat-value">${combinedNewStudents.length}</div>
            </div>
            <div class="stat-item">
                <div class="stat-label">仅在A表存在</div>
                <div class="stat-value">${onlyInA}</div>
            </div>
        </div>
    `;
    
    differencesTableContainer.appendChild(summaryDiv);

    // 按工作表分组数据
    const groupedBySheet = allSheetResults.bySheet;
    
    // 创建工作表导航选项卡
    const tabsContainer = document.createElement('div');
    tabsContainer.className = 'sheet-tabs-container';
    
    const tabsList = document.createElement('ul');
    tabsList.className = 'sheet-tabs';
    
    // 添加"全部"选项卡
    const allTab = document.createElement('li');
    allTab.className = 'sheet-tab active';
    allTab.dataset.sheetId = 'all';
    allTab.textContent = `全部 (${totalItems})`;
    tabsList.appendChild(allTab);
    
    // 为全部标签添加点击事件
    allTab.addEventListener('click', function() {
        switchToTab('all');
    });
    
    // 添加每个工作表的选项卡
    Object.keys(groupedBySheet).forEach(sheetName => {
        const sheetData = groupedBySheet[sheetName];
        const sheetItemCount = sheetData.differences.length + sheetData.newStudents.length;
        
        // 如果没有任何项，跳过这个工作表
        if (sheetItemCount === 0) return;
        if (sheetItemCount > 0) {
            const tab = document.createElement('li');
            tab.className = 'sheet-tab';
            tab.dataset.sheetId = sheetName;
            tab.textContent = `${sheetName} (${sheetItemCount})`;
            
            // 添加点击事件处理
            tab.addEventListener('click', function() {
                switchToTab(sheetName);
            });
            
            tabsList.appendChild(tab);
        }
    });
    
    tabsContainer.appendChild(tabsList);
    differencesTableContainer.appendChild(tabsContainer);
    
    // 创建工作表内容容器
    const sheetsContentContainer = document.createElement('div');
    sheetsContentContainer.className = 'sheets-content-container';
    differencesTableContainer.appendChild(sheetsContentContainer);
    
    // 创建"全部"视图的内容
    const allContent = createSheetContentView('all', combinedDifferences, combinedNewStudents);
    allContent.classList.add('active');
    sheetsContentContainer.appendChild(allContent);
    
    // 创建每个工作表的内容视图
    Object.keys(groupedBySheet).forEach(sheetName => {
        const sheetData = groupedBySheet[sheetName];
        if (sheetData.differences.length > 0 || sheetData.newStudents.length > 0) {
            const sheetContent = createSheetContentView(
                sheetName, 
                sheetData.differences, 
                sheetData.newStudents
            );
            sheetsContentContainer.appendChild(sheetContent);
        }
    });
    
    // 添加选项卡点击事件
    tabsList.querySelectorAll('.sheet-tab').forEach(tab => {
        tab.addEventListener('click', () => {
            // 切换选项卡激活状态
            tabsList.querySelectorAll('.sheet-tab').forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            
            // 切换内容视图
            const sheetId = tab.dataset.sheetId;
            sheetsContentContainer.querySelectorAll('.sheet-content').forEach(content => {
                if (content.dataset.sheetId === sheetId) {
                    content.classList.add('active');
                } else {
                    content.classList.remove('active');
                }
            });
        });
    });
    
    // 显示结果容器
    differencesContainer.classList.remove('hidden');
    
    // 显示快速更新按钮
    quickUpdateBtn.classList.remove('hidden');
    
    // 添加快速更新按钮点击事件
    quickUpdateBtn.addEventListener('click', scrollToUpdateButton);
}

/**
 * 生成包含工作表详细信息的结果摘要
 * @param {Object} allSheetResults - 所有工作表的比对结果
 * @param {number} totalA - A表总记录数
 * @param {number} totalB - B表总记录数
 */
function generateResultsSummaryWithSheets(allSheetResults, totalA, totalB) {
    // 显示摘要容器
    summaryDiv.classList.remove('hidden');
    
    // 统计总数据
    const totalDifferences = allSheetResults.combined.differences.length;
    const totalNewStudents = allSheetResults.combined.newStudents.length;
    
    // 获取工作表统计信息
    const sheetStats = allSheetResults.sheetStats || [];
    
    // 构建工作表摘要HTML
    let sheetsSummaryHTML = '';
    if (sheetStats.length > 0) {
        sheetsSummaryHTML = `
            <h4>工作表比对结果：</h4>
            <ul>
                ${sheetStats.map(sheet => 
                    `<li><strong>${sheet.name}</strong>: A表${sheet.recordsA}条记录, B表${sheet.recordsB}条记录, 
                     需更新${sheet.differences}条, 新增${sheet.newStudents}条</li>`
                ).join('')}
            </ul>
        `;
    }

    // 构建摘要HTML
    summaryContent.innerHTML = `
        <div>
            <p><strong>表格A总记录数：</strong> ${totalA}条（来自${sheetsInfoA.selectedSheets.size}个工作表）</p>
            <p><strong>表格B总记录数：</strong> ${totalB}条（来自${sheetsInfoB.selectedSheets.size}个工作表）</p>
            <p><strong>比对结果：</strong></p>
            <ul>
                <li class="highlight-diff"><strong>需要更新的记录: ${totalDifferences}条</strong></li>
                <li class="highlight-diff"><strong>新增学生: ${totalNewStudents}条</strong></li>
            </ul>
            ${sheetsSummaryHTML}
            <p><em>提示：可以在下方的工作表选项卡中切换查看各工作表的详细结果</em></p>
        </div>
    `;
}

/**
 * 切换到指定工作表的选项卡
 * @param {string} tabId - 要切换到的工作表ID
 */
function switchToTab(tabId) {
    console.log(`切换到工作表选项卡: ${tabId}`);
    
    // 获取所有选项卡和内容
    const tabs = document.querySelectorAll('.sheet-tab');
    const contents = document.querySelectorAll('.sheet-content');
    
    // 先隐藏所有内容并移除选项卡的active类
    tabs.forEach(tab => tab.classList.remove('active'));
    contents.forEach(content => content.classList.remove('active'));
    
    // 激活选中的选项卡和内容
    // 使用两种方式查找元素，以提高兼容性
    const selectedTab = document.querySelector(`.sheet-tab[data-sheet-id="${tabId}"]`) || 
                        document.querySelector(`.sheet-tab[dataset-sheet-id="${tabId}"]`) || 
                        document.querySelector(`.sheet-tab[data-sheetid="${tabId}"]`);
                        
    const selectedContent = document.querySelector(`.sheet-content[data-sheet-id="${tabId}"]`) || 
                            document.querySelector(`.sheet-content[dataset-sheet-id="${tabId}"]`) || 
                            document.querySelector(`.sheet-content[data-sheetid="${tabId}"]`);
    
    if (selectedTab) {
        selectedTab.classList.add('active');
        console.log(`选项卡 ${tabId} 已激活`);
    } else {
        console.warn(`未找到ID为${tabId}的选项卡`);
    }
    
    if (selectedContent) {
        selectedContent.classList.add('active');
        console.log(`内容 ${tabId} 已显示`);
    } else {
        console.warn(`未找到ID为${tabId}的内容`);
    }
}

// 初始化搜索和筛选功能
document.addEventListener('DOMContentLoaded', function() {
    // 初始化搜索功能
    const searchInput = document.getElementById('searchInput');
    if (searchInput) {
        searchInput.addEventListener('input', filterResults);
    }
    
    // 初始化班级筛选下拉菜单
    const classFilter = document.getElementById('classFilter');
    if (classFilter) {
        classFilter.addEventListener('change', filterResults);
    }
});

/**
 * 根据搜索条件和筛选条件过滤结果
 */
function filterResults() {
    const searchInput = document.getElementById('searchInput');
    const classFilter = document.getElementById('classFilter');
    
    if (!searchInput || !classFilter) return;
    
    const searchTerm = searchInput.value.toLowerCase().trim();
    const selectedClass = classFilter.value;
    
    // 获取所有差异行
    const rows = document.querySelectorAll('#differences-table-container tr[data-type]');
    if (!rows.length) return;
    
    // 遍历所有行并应用筛选
    rows.forEach(row => {
        // 默认隐藏所有行
        row.style.display = 'none';
        
        // 获取姓名单元格内容
        const nameCell = row.querySelector('td:nth-child(3)');
        const classCell = row.querySelector('td:nth-child(4)');
        
        if (!nameCell || !classCell) return;
        
        const name = nameCell.textContent.toLowerCase();
        const className = classCell.textContent;
        
        // 检查是否符合搜索和筛选条件
        const matchesSearch = !searchTerm || name.includes(searchTerm);
        const matchesClass = selectedClass === 'all' || className === selectedClass;
        
        // 不再考虑showIdentical选项，总是显示结果
        // 如果同时满足搜索和班级筛选条件，显示该行
        if (matchesSearch && matchesClass) {
            row.style.display = '';
            
            // 如果是分组的第一行，确保显示整个分组
            if (row.classList.contains('row-group') && row.dataset.type === 'update') {
                let nextRow = row.nextElementSibling;
                while (nextRow && nextRow.classList.contains('row-group-child')) {
                    nextRow.style.display = '';
                    nextRow = nextRow.nextElementSibling;
                }
            }
        }
    });
    
    // 更新统计数字
    updateFilterStats();
}

/**
 * 更新筛选后的统计数字
 */
function updateFilterStats() {
    const visibleUpdateRows = document.querySelectorAll('#differences-table-container tr[data-type="update"]:not([style*="display: none"])').length;
    const visibleNewRows = document.querySelectorAll('#differences-table-container tr[data-type="new"]:not([style*="display: none"])').length;
    const totalVisible = visibleUpdateRows + visibleNewRows;
    
    // 更新统计卡片中的数字
    const updateCount = document.querySelector('.stat-card:nth-child(2) .stat-card-value');
    const newCount = document.querySelector('.stat-card:nth-child(3) .stat-card-value');
    const totalCount = document.querySelector('.stat-card:nth-child(1) .stat-card-value');
    
    if (updateCount) updateCount.textContent = visibleUpdateRows;
    if (newCount) newCount.textContent = visibleNewRows;
    if (totalCount) totalCount.textContent = totalVisible;
    
    // 更新标签页中的数字
    const updateTab = document.querySelector('.segment-tab[data-filter="update"] .segment-tab-count');
    const newTab = document.querySelector('.segment-tab[data-filter="new"] .segment-tab-count');
    const allTab = document.querySelector('.segment-tab[data-filter="all"] .segment-tab-count');
    
    if (updateTab) updateTab.textContent = visibleUpdateRows;
    if (newTab) newTab.textContent = visibleNewRows;
    if (allTab) allTab.textContent = totalVisible;
}

/**
 * 从比对结果中提取班级列表并填充筛选下拉框
 */
function populateClassFilter() {
    const classFilter = document.getElementById('classFilter');
    if (!classFilter) return;
    
    // 清空现有选项（保留"所有班级"选项）
    while (classFilter.options.length > 1) {
        classFilter.remove(1);
    }
    
    // 获取所有班级单元格
    const classCells = document.querySelectorAll('#differences-table-container td:nth-child(4)');
    const classSet = new Set();
    
    // 收集所有唯一的班级名称
    classCells.forEach(cell => {
        const className = cell.textContent.trim();
        if (className) {
            classSet.add(className);
        }
    });
    
    // 按字母顺序排序班级
    const sortedClasses = Array.from(classSet).sort();
    
    // 填充下拉框
    sortedClasses.forEach(className => {
        const option = document.createElement('option');
        option.value = className;
        option.textContent = className;
        classFilter.appendChild(option);
    });
}

// 在显示比对结果后调用此函数
const originalDisplayDifferences = displayDifferences;
displayDifferences = function(data) {
    originalDisplayDifferences(data);
    // 填充班级筛选下拉框
    populateClassFilter();
};

// 页面加载完成后运行
document.addEventListener('DOMContentLoaded', function() {
    // 初始化导航切换功能
    initNavigation();
    
    // 初始化对比功能
    initCompareFeature();
    
    // 初始化统计功能
    initStatsFeature();
});

/**
 * 初始化导航切换功能
 */
function initNavigation() {
    const navTabs = document.querySelectorAll('.nav-tab');
    const appPages = document.querySelectorAll('.app-page');
    
    navTabs.forEach(tab => {
        tab.addEventListener('click', () => {
            // 移除所有激活状态
            navTabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            
            // 隐藏所有页面
            appPages.forEach(page => page.classList.remove('active'));
            
            // 显示目标页面
            const targetPageId = tab.getAttribute('data-page');
            document.getElementById(targetPageId).classList.add('active');
        });
    });
}

/**
 * 初始化对比功能
 */
function initCompareFeature() {
    // 文件A上传处理
    const fileA_input = document.getElementById('fileA');
    if (fileA_input) {
        fileA_input.addEventListener('change', function(event) {
            handleFileUpload(event, 'A');
        });
    }
    
    // 文件B上传处理
    const fileB_input = document.getElementById('fileB');
    if (fileB_input) {
        fileB_input.addEventListener('change', function(event) {
            handleFileUpload(event, 'B');
        });
    }
    
    // 比对按钮点击事件
    const compareBtn = document.getElementById('compareBtn');
    if (compareBtn) {
        compareBtn.addEventListener('click', handleCompare);
    }
    
    // 确认更新按钮点击事件
    const confirmUpdateBtn = document.getElementById('confirmUpdateBtn');
    if (confirmUpdateBtn) {
        confirmUpdateBtn.addEventListener('click', handleConfirmUpdate);
    }
    
    // 撤销更新按钮点击事件
    const undoUpdateBtn = document.getElementById('undoUpdateBtn');
    if (undoUpdateBtn) {
        undoUpdateBtn.addEventListener('click', handleUndoUpdate);
    }
    
    // 快速更新按钮点击事件
    const quickUpdateBtn = document.getElementById('quick-update-btn');
    if (quickUpdateBtn) {
        quickUpdateBtn.addEventListener('click', scrollToUpdateButton);
    }
    
    // 初始化全选按钮
    const selectAllA = document.getElementById('selectAllA');
    if (selectAllA) {
        selectAllA.addEventListener('click', () => toggleAllSheets('A'));
    }
    
    const selectAllB = document.getElementById('selectAllB');
    if (selectAllB) {
        selectAllB.addEventListener('click', () => toggleAllSheets('B'));
    }
    
    // 清除筛选按钮点击事件
    const clearCompareFiltersBtn = document.getElementById('clearCompareFilters');
    if (clearCompareFiltersBtn) {
        clearCompareFiltersBtn.addEventListener('click', () => {
            clearFilters('compare-filters');
        });
    }
}

/**
 * 初始化统计功能
 */
function initStatsFeature() {
    // 统计文件上传处理
    const fileStats_input = document.getElementById('fileStats');
    if (fileStats_input) {
        fileStats_input.addEventListener('change', function(event) {
            handleStatsFileUpload(event);
        });
    }
    
    // 初始化模式选择器
    initModeSelector();
    
    // 生成统计按钮点击事件
    const generateStatsBtn = document.getElementById('generateStatsBtn');
    if (generateStatsBtn) {
        generateStatsBtn.addEventListener('click', generateStatistics);
    }
    
    // 导出统计报表按钮点击事件
    const exportStatsBtn = document.getElementById('exportStatsBtn');
    if (exportStatsBtn) {
        exportStatsBtn.addEventListener('click', exportStatistics);
    }
    
    // 清除筛选按钮点击事件
    const clearStatsFiltersBtn = document.getElementById('clearStatsFilters');
    if (clearStatsFiltersBtn) {
        clearStatsFiltersBtn.addEventListener('click', () => {
            clearFilters('stats-filters');
        });
    }
    
    // 模板保存和加载按钮
    const saveTemplateBtn = document.getElementById('saveTemplateBtn');
    if (saveTemplateBtn) {
        saveTemplateBtn.addEventListener('click', saveStatsTemplate);
    }
    
    const loadTemplateBtn = document.getElementById('loadTemplateBtn');
    if (loadTemplateBtn) {
        loadTemplateBtn.addEventListener('click', loadStatsTemplate);
    }
}

/**
 * 清除筛选条件
 * @param {string} containerId - 筛选容器的ID
 */
function clearFilters(containerId) {
    const filterContainer = document.getElementById(containerId);
    if (!filterContainer) return;
    
    // 重置所有下拉框为"全部"选项
    const selects = filterContainer.querySelectorAll('.filter-select');
    selects.forEach(select => {
        select.value = 'all';
    });
    
    // 触发筛选应用
    applyFilters();
    
    // 更新状态提示
    const filterStatusEl = document.querySelector('.filter-status');
    if (filterStatusEl) {
        filterStatusEl.classList.add('hidden');
    }
}

// ===== 统计功能相关函数 =====

// 统计数据存储
let statsData = null;
let statsWorkbook = null;

/**
 * 处理统计文件上传
 * @param {Event} event - 文件上传事件
 */
function handleStatsFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    // 更新上传区域显示
    const uploadArea = document.getElementById('upload-area-stats');
    const fileInfo = uploadArea.querySelector('.file-info');
    const fileName = uploadArea.querySelector('.file-name');
    
    // 显示文件名
    if (fileInfo && fileName) {
        uploadArea.classList.add('file-uploaded');
        fileInfo.classList.remove('hidden');
        fileName.textContent = file.name;
    }
    
    // 显示状态消息
    const statusDiv = document.getElementById('stats-status');
    if (statusDiv) {
        statusDiv.textContent = '文件已上传，请点击"生成统计报表"按钮进行分析';
        statusDiv.style.color = 'var(--primary)';
    }
}

/**
 * 生成统计数据和图表
 */
async function generateStatistics() {
    const statusDiv = document.getElementById('stats-status');
    
    // 检查文件是否已上传
    const fileInput = document.getElementById('fileStats');
    if (!fileInput || !fileInput.files[0]) {
        if (statusDiv) {
            statusDiv.textContent = '请先上传Excel文件！';
            statusDiv.style.color = 'var(--danger)';
        }
        return;
    }
    
    // 检查XLSX库是否已加载
    if (typeof XLSX === 'undefined') {
        if (statusDiv) {
            statusDiv.textContent = '错误：Excel处理库(XLSX)未加载，正在尝试重新加载...';
            statusDiv.style.color = 'var(--danger)';
        }
        loadXLSXLibrary();
        return;
    }
    
    if (statusDiv) {
        statusDiv.textContent = '正在处理数据，请稍候...';
        statusDiv.style.color = 'var(--primary)';
    }
    
    try {
        // 读取Excel文件
        const file = fileInput.files[0];
        statsWorkbook = await readExcelFile(file);
        
        // 提取数据
        const allData = [];
        statsWorkbook.SheetNames.forEach(sheetName => {
            const worksheet = statsWorkbook.Sheets[sheetName];
            const sheetData = XLSX.utils.sheet_to_json(worksheet);
            
            // 为每条记录添加工作表标识
            const dataWithSheetInfo = sheetData.map(row => ({
                ...row,
                '_原始工作表': sheetName
            }));
            
            allData.push(...dataWithSheetInfo);
        });
        
        // 存储数据
        statsData = allData;
        
        // 检查是否有数据
        if (allData.length === 0) {
            if (statusDiv) {
                statusDiv.textContent = '错误：Excel文件中没有找到数据！';
                statusDiv.style.color = 'var(--danger)';
            }
            return;
        }
        
        // 模板模式处理
        if (currentStatsMode === 'template') {
            // 先尝试加载模板文件（如果尚未加载）
            if (!statsTemplateWorkbook) {
                // 尝试加载模板文件
                try {
                    await loadTemplateFile();
                } catch (e) {
                    // 加载失败时提示用户
                    console.warn('模板文件加载失败');
                    if (statusDiv) {
                        statusDiv.textContent = '未找到可用的模板文件，请确保"统计"文件夹中有Excel模板文件';
                        statusDiv.style.color = 'var(--warning)';
                    }
                    return;
                }
            }
            
            // 如果有模板文件，使用模板文件生成统计
            if (statsTemplateWorkbook) {
                if (statusDiv) {
                    statusDiv.textContent = '正在按模板生成统计报表...';
                }
                
                // 使用模板文件生成统计
                await generateStatsByTemplate(allData);
                return;
            } else {
                // 没有找到模板文件，显示错误信息
                if (statusDiv) {
                    statusDiv.textContent = '请先选择一个模板文件';
                    statusDiv.style.color = 'var(--danger)';
                }
                return;
            }
        }
        
        // 根据不同模式分析数据
        let statistics;
        if (currentStatsMode === 'normal') {
            // 普通模式：自动检测字段
            statistics = analyzeStatsData(allData);
        } else {
            // 模板模式：根据选中的字段生成统计
            statistics = analyzeTemplateData(allData);
        }
        
        // 显示统计结果
        displayStatistics(statistics);
        
        if (statusDiv) {
            statusDiv.textContent = `统计分析完成，共处理了${allData.length}条记录`;
            statusDiv.style.color = 'var(--success)';
        }
        
        // 显示结果容器
        const resultsContainer = document.getElementById('stats-results-container');
        if (resultsContainer) {
            resultsContainer.classList.remove('hidden');
        }
        
    } catch (error) {
        console.error('统计分析错误:', error);
        if (statusDiv) {
            statusDiv.textContent = `统计分析出错: ${error.message}`;
            statusDiv.style.color = 'var(--danger)';
        }
    }
}

/**
 * 分析统计数据
 * @param {Array} data - 要分析的数据
 * @returns {Object} - 分析结果
 */
function analyzeStatsData(data) {
    const result = {
        totalRecords: data.length,
        sheets: {},
        gender: {},
        class: {},
        grade: {},
        score: {}
    };
    
    // 分析工作表分布
    data.forEach(row => {
        const sheet = row['_原始工作表'] || '未知';
        if (!result.sheets[sheet]) {
            result.sheets[sheet] = 0;
        }
        result.sheets[sheet]++;
    });
    
    // 寻找可能的性别字段
    const possibleGenderFields = ['性别', '性别代码', 'gender', 'sex'];
    let genderField = null;
    
    for (const field of possibleGenderFields) {
        if (data[0] && data[0][field] !== undefined) {
            genderField = field;
            break;
        }
    }
    
    if (genderField) {
        data.forEach(row => {
            const gender = row[genderField] ? String(row[genderField]).trim() : '未知';
            if (!result.gender[gender]) {
                result.gender[gender] = 0;
            }
            result.gender[gender]++;
        });
    }
    
    // 寻找可能的班级字段
    const possibleClassFields = ['班级', '班号', 'class'];
    let classField = null;
    
    for (const field of possibleClassFields) {
        if (data[0] && data[0][field] !== undefined) {
            classField = field;
            break;
        }
    }
    
    if (classField) {
        data.forEach(row => {
            const className = row[classField] ? String(row[classField]).trim() : '未知';
            if (!result.class[className]) {
                result.class[className] = 0;
            }
            result.class[className]++;
        });
    }
    
    // 寻找可能的年级字段
    const possibleGradeFields = ['年级', 'grade'];
    let gradeField = null;
    
    for (const field of possibleGradeFields) {
        if (data[0] && data[0][field] !== undefined) {
            gradeField = field;
            break;
        }
    }
    
    if (gradeField) {
        data.forEach(row => {
            const grade = row[gradeField] ? String(row[gradeField]).trim() : '未知';
            if (!result.grade[grade]) {
                result.grade[grade] = 0;
            }
            result.grade[grade]++;
        });
    }
    
    // 寻找可能的成绩字段（可能有多个）
    const scoreFields = [];
    const possibleScoreFields = ['成绩', '分数', '总分', '语文', '数学', '英语', 'score'];
    
    // 检查每个可能的字段是否存在
    for (const field of possibleScoreFields) {
        if (data[0] && data[0][field] !== undefined && !isNaN(parseFloat(data[0][field]))) {
            scoreFields.push(field);
        }
    }
    
    // 如果找到了成绩字段
    if (scoreFields.length > 0) {
        // 初始化成绩区间
        const scoreRanges = {
            '0-59': 0,
            '60-69': 0,
            '70-79': 0,
            '80-89': 0,
            '90-100': 0
        };
        
        // 对每个成绩字段进行统计
        scoreFields.forEach(field => {
            result.score[field] = JSON.parse(JSON.stringify(scoreRanges)); // 深拷贝
            
            data.forEach(row => {
                const score = parseFloat(row[field]);
                if (!isNaN(score)) {
                    if (score < 60) result.score[field]['0-59']++;
                    else if (score < 70) result.score[field]['60-69']++;
                    else if (score < 80) result.score[field]['70-79']++;
                    else if (score < 90) result.score[field]['80-89']++;
                    else result.score[field]['90-100']++;
                }
            });
        });
    }
    
    return result;
}

/**
 * 显示统计结果
 * @param {Object} statistics - 统计分析结果
 */
function displayStatistics(statistics) {
    // 创建统计表格
    displayStatsTable(statistics);
}

/**
 * 显示统计表格
 * @param {Object} statistics - 统计数据
 */
function displayStatsTable(statistics) {
    const tableContainer = document.getElementById('stats-table-container');
    if (!tableContainer) return;
    
    let tableHtml = '<table class="stats-table">';
    tableHtml += '<thead><tr><th>统计项目</th><th>统计值</th><th>数量</th><th>百分比</th></tr></thead>';
    tableHtml += '<tbody>';
    
    // 工作表分布
    tableHtml += '<tr class="table-section"><td colspan="4">工作表分布</td></tr>';
    Object.entries(statistics.sheets).forEach(([sheet, count]) => {
        const percent = Math.round((count / statistics.totalRecords) * 100);
        tableHtml += `<tr><td>工作表</td><td>${sheet}</td><td>${count}</td><td>${percent}%</td></tr>`;
    });
    
    // 根据当前模式显示不同的统计结果
    if (currentStatsMode === 'normal') {
        // 普通模式：显示自动检测的统计结果
        
        // 性别分布
        if (statistics.gender && Object.keys(statistics.gender).length > 0) {
            tableHtml += '<tr class="table-section"><td colspan="4">性别分布</td></tr>';
            const genderTotal = Object.values(statistics.gender).reduce((a, b) => a + b, 0);
            Object.entries(statistics.gender).forEach(([gender, count]) => {
                const percent = Math.round((count / genderTotal) * 100);
                tableHtml += `<tr><td>性别</td><td>${gender}</td><td>${count}</td><td>${percent}%</td></tr>`;
            });
        }
        
        // 班级分布
        if (statistics.class && Object.keys(statistics.class).length > 0) {
            tableHtml += '<tr class="table-section"><td colspan="4">班级分布</td></tr>';
            const classTotal = Object.values(statistics.class).reduce((a, b) => a + b, 0);
            Object.entries(statistics.class).forEach(([className, count]) => {
                const percent = Math.round((count / classTotal) * 100);
                tableHtml += `<tr><td>班级</td><td>${className}</td><td>${count}</td><td>${percent}%</td></tr>`;
            });
        }
        
        // 年级分布
        if (statistics.grade && Object.keys(statistics.grade).length > 0) {
            tableHtml += '<tr class="table-section"><td colspan="4">年级分布</td></tr>';
            const gradeTotal = Object.values(statistics.grade).reduce((a, b) => a + b, 0);
            Object.entries(statistics.grade).forEach(([grade, count]) => {
                const percent = Math.round((count / gradeTotal) * 100);
                tableHtml += `<tr><td>年级</td><td>${grade}</td><td>${count}</td><td>${percent}%</td></tr>`;
            });
        }
        
        // 成绩分布
        if (statistics.score && Object.keys(statistics.score).length > 0) {
            tableHtml += '<tr class="table-section"><td colspan="4">成绩分布</td></tr>';
            Object.entries(statistics.score).forEach(([subject, ranges]) => {
                const subjectTotal = Object.values(ranges).reduce((a, b) => a + b, 0);
                Object.entries(ranges).forEach(([range, count]) => {
                    const percent = Math.round((count / subjectTotal) * 100);
                    tableHtml += `<tr><td>${subject}</td><td>${range}分</td><td>${count}</td><td>${percent}%</td></tr>`;
                });
            });
        }
        
    } else {
        // 模板模式：显示自定义统计结果
        if (statistics.custom && Object.keys(statistics.custom).length > 0) {
            // 遍历每个自定义字段的统计结果
            Object.entries(statistics.custom).forEach(([field, values]) => {
                // 排除分数区间统计，它们会在字段统计后单独显示
                if (field.endsWith('_分数区间')) return;
                
                tableHtml += `<tr class="table-section"><td colspan="4">${field}分布</td></tr>`;
                
                const fieldTotal = Object.values(values).reduce((a, b) => a + b, 0);
                
                // 按数量降序排序
                const sortedValues = Object.entries(values)
                    .sort((a, b) => b[1] - a[1]);
                    
                sortedValues.forEach(([value, count]) => {
                    const percent = Math.round((count / fieldTotal) * 100);
                    tableHtml += `<tr><td>${field}</td><td>${value}</td><td>${count}</td><td>${percent}%</td></tr>`;
                });
                
                // 检查是否有对应的分数区间统计
                const scoreRangeField = `${field}_分数区间`;
                if (statistics.custom[scoreRangeField]) {
                    tableHtml += `<tr class="table-section"><td colspan="4">${field}分数区间</td></tr>`;
                    
                    const rangeTotal = Object.values(statistics.custom[scoreRangeField]).reduce((a, b) => a + b, 0);
                    
                    // 确保分数区间按正确顺序排序
                    const rangeOrder = ['0-59', '60-69', '70-79', '80-89', '90-100', '其他'];
                    const sortedRanges = Object.entries(statistics.custom[scoreRangeField])
                        .sort((a, b) => rangeOrder.indexOf(a[0]) - rangeOrder.indexOf(b[0]));
                    
                    sortedRanges.forEach(([range, count]) => {
                        const percent = Math.round((count / rangeTotal) * 100);
                        const rangeDisplay = range === '其他' ? range : `${range}分`;
                        tableHtml += `<tr><td>${field}分数</td><td>${rangeDisplay}</td><td>${count}</td><td>${percent}%</td></tr>`;
                    });
                }
            });
        } else {
            tableHtml += '<tr><td colspan="4">未配置统计字段或数据为空</td></tr>';
        }
    }
    
    tableHtml += '</tbody></table>';
    
    tableContainer.innerHTML = tableHtml;
}

/**
 * 导出统计数据为Excel
 */
function exportStatistics() {
    if (!statsData || statsData.length === 0) {
        alert('没有可导出的数据！');
        return;
    }
    
    // 创建工作簿
    const wb = XLSX.utils.book_new();
    
    // 创建原始数据工作表
    const ws_data = XLSX.utils.json_to_sheet(statsData);
    XLSX.utils.book_append_sheet(wb, ws_data, "原始数据");
    
    // 获取当前显示的统计结果
    const tableContainer = document.getElementById('stats-table-container');
    if (tableContainer && tableContainer.querySelector('table')) {
        try {
            // 从HTML表格转换为工作表
            const ws_stats = XLSX.utils.table_to_sheet(tableContainer.querySelector('table'));
            
            // 添加统计结果工作表
            const sheetName = currentStatsMode === 'normal' ? "自动统计结果" : "模板统计结果";
            XLSX.utils.book_append_sheet(wb, ws_stats, sheetName);
        } catch (e) {
            console.error('转换统计表格失败:', e);
        }
    }
    
    // 如果是模板模式，添加模板配置工作表
    if (currentStatsMode === 'template' && statsTemplateConfig) {
        try {
            // 将模板配置转换为工作表数据
            const templateData = [
                { 名称: '模板名称', 值: statsTemplateConfig.name },
                { 名称: '创建时间', 值: new Date(statsTemplateConfig.createdAt).toLocaleString() },
                { 名称: '字段数量', 值: statsTemplateConfig.fields.length }
            ];
            
            // 添加字段列表
            statsTemplateConfig.fields.forEach((field, index) => {
                templateData.push({ 名称: `字段${index + 1}`, 值: field });
            });
            
            const ws_template = XLSX.utils.json_to_sheet(templateData);
            XLSX.utils.book_append_sheet(wb, ws_template, "模板配置");
        } catch (e) {
            console.error('添加模板配置失败:', e);
        }
    }
    
    try {
        // 获取当前日期时间作为文件名
        const now = new Date();
        const timestamp = now.getFullYear().toString() +
            ('0' + (now.getMonth() + 1)).slice(-2) +
            ('0' + now.getDate()).slice(-2) +
            ('0' + now.getHours()).slice(-2) +
            ('0' + now.getMinutes()).slice(-2);
        
        // 文件名包含当前模式
        const modeText = currentStatsMode === 'normal' ? '普通模式' : '模板模式';
        
        // 导出文件
        XLSX.writeFile(wb, `数据统计_${modeText}_${timestamp}.xlsx`);
        
        // 显示成功消息
        const statusDiv = document.getElementById('stats-status');
        if (statusDiv) {
            statusDiv.textContent = '统计报表已导出';
            statusDiv.style.color = 'var(--success)';
        }
    } catch (e) {
        console.error('导出统计报表失败:', e);
        alert('导出统计报表失败: ' + e.message);
    }
}

/**
 * 自动检测表格中可筛选的字段
 * @param {Array} data - Excel表格数据
 * @returns {Object} 可筛选的字段及其唯一值
 */
function detectFilterFields(data) {
    if (!data || !data.length || data.length === 0) return {};
    
    // 获取表格的所有列名
    const columns = Object.keys(data[0]);
    const filterFields = {};
    
    // 可能的筛选字段关键词
    const possibleFilterFields = {
        school: ['学校', '校名', 'school'],
        grade: ['年级', 'grade'],
        class: ['班级', '班号', 'class'],
        gender: ['性别', 'gender', 'sex']
    };
    
    // 遍历所有列，检查是否为可筛选字段
    columns.forEach(column => {
        // 检查列名是否匹配筛选字段关键词
        let fieldType = null;
        for (const [type, keywords] of Object.entries(possibleFilterFields)) {
            if (keywords.some(keyword => column.includes(keyword))) {
                fieldType = type;
                break;
            }
        }
        
        // 如果不是预设的筛选字段类型，检查列中的唯一值数量
        if (!fieldType) {
            const uniqueValues = new Set();
            data.forEach(row => {
                if (row[column] !== undefined) {
                    uniqueValues.add(String(row[column]).trim());
                }
            });
            
            // 如果唯一值数量在2-20之间，可能是可筛选字段
            if (uniqueValues.size >= 2 && uniqueValues.size <= 20) {
                fieldType = 'custom';
            } else {
                return; // 不是可筛选字段
            }
        }
        
        // 收集字段的所有唯一值
        const uniqueValues = new Set();
        data.forEach(row => {
            if (row[column] !== undefined) {
                uniqueValues.add(String(row[column]).trim());
            }
        });
        
        // 转换为数组并排序
        const sortedValues = Array.from(uniqueValues).sort();
        
        // 将字段及其唯一值添加到结果中
        filterFields[column] = {
            type: fieldType,
            values: sortedValues
        };
    });
    
    return filterFields;
}

/**
 * 生成筛选界面
 * @param {Object} filterFields - 可筛选的字段及其唯一值
 * @param {string} containerId - 筛选容器的ID
 */
function generateFilterUI(filterFields, containerId) {
    const filterContainer = document.getElementById(containerId);
    if (!filterContainer) return;
    
    filterContainer.innerHTML = '';
    
    // 对筛选字段进行排序处理，优先显示学校、年级、班级
    const priorityOrder = ['school', 'grade', 'class', 'gender'];
    
    // 按优先级和字段名排序
    const sortedFields = Object.entries(filterFields).sort((a, b) => {
        const typeA = a[1].type;
        const typeB = b[1].type;
        
        // 优先按预设字段类型排序
        const priorityA = priorityOrder.indexOf(typeA);
        const priorityB = priorityOrder.indexOf(typeB);
        
        if (priorityA !== -1 && priorityB !== -1) {
            return priorityA - priorityB;
        } else if (priorityA !== -1) {
            return -1;
        } else if (priorityB !== -1) {
            return 1;
        }
        
        // 如果都不是预设字段或都是同一类型，按字段名排序
        return a[0].localeCompare(b[0]);
    });
    
    // 创建筛选界面
    for (const [fieldName, fieldInfo] of sortedFields) {
        const values = fieldInfo.values;
        
        // 创建筛选选择器
        const filterOption = document.createElement('div');
        filterOption.className = 'filter-option';
        
        // 创建标签
        const label = document.createElement('label');
        label.textContent = fieldName + ':';
        label.className = 'filter-label';
        
        // 创建下拉框
        const select = document.createElement('select');
        select.className = 'filter-select';
        select.dataset.field = fieldName;
        
        // 添加"全部"选项
        const allOption = document.createElement('option');
        allOption.value = 'all';
        allOption.textContent = `全部${fieldName}`;
        select.appendChild(allOption);
        
        // 添加字段的唯一值作为选项
        values.forEach(value => {
            const option = document.createElement('option');
            option.value = value;
            option.textContent = value;
            select.appendChild(option);
        });
        
        // 监听选择事件
        select.addEventListener('change', function() {
            applyFilters();
        });
        
        // 将标签和下拉框添加到筛选选项
        filterOption.appendChild(label);
        filterOption.appendChild(select);
        
        // 将筛选选项添加到容器
        filterContainer.appendChild(filterOption);
    }
}

/**
 * 应用筛选条件
 */
function applyFilters() {
    // 获取所有筛选选择器
    const filterSelects = document.querySelectorAll('.filter-select');
    const activeFilters = {};
    
    // 收集活跃的筛选条件
    filterSelects.forEach(select => {
        const field = select.dataset.field;
        const value = select.value;
        
        if (value !== 'all') {
            activeFilters[field] = value;
        }
    });
    
    // 根据当前页面确定要筛选的数据
    let dataToFilter = [];
    let displayFunction = null;
    
    // 判断当前激活的页面
    const activePageId = document.querySelector('.app-page.active').id;
    
    if (activePageId === 'compare-page') {
        // 对比页面的筛选处理
        if (globalComparisonResult && globalComparisonResult.differences) {
            dataToFilter = globalComparisonResult.differences;
            displayFunction = function(filteredData) {
                const filteredResult = {
                    ...globalComparisonResult,
                    differences: filteredData
                };
                displayDifferences(filteredResult);
            };
        }
    } else if (activePageId === 'stats-page') {
        // 统计页面的筛选处理
        if (statsData && statsData.length > 0) {
            dataToFilter = statsData;
            displayFunction = function(filteredData) {
                const statistics = analyzeStatsData(filteredData);
                displayStatistics(statistics);
                
                // 更新统计计数
                const statusDiv = document.getElementById('stats-status');
                if (statusDiv) {
                    statusDiv.textContent = `统计分析显示: 共${filteredData.length}/${statsData.length}条记录`;
                    statusDiv.style.color = 'var(--primary)';
                }
            };
        }
    }
    
    // 如果没有数据可筛选，直接返回
    if (dataToFilter.length === 0 || !displayFunction) return;
    
    // 应用筛选条件
    const filteredData = dataToFilter.filter(item => {
        // 检查是否满足所有筛选条件
        return Object.entries(activeFilters).every(([field, value]) => {
            // 处理可能的undefined值
            const itemValue = item[field] !== undefined ? String(item[field]).trim() : '';
            return itemValue === value;
        });
    });
    
    // 显示筛选后的数据
    displayFunction(filteredData);
    
    // 更新筛选状态信息
    updateFilterStatus(Object.keys(activeFilters).length, filteredData.length, dataToFilter.length);
}

/**
 * 更新筛选状态信息
 * @param {number} filterCount - 活跃的筛选条件数量
 * @param {number} resultCount - 筛选后的记录数
 * @param {number} totalCount - 总记录数
 */
function updateFilterStatus(filterCount, resultCount, totalCount) {
    const filterStatusEl = document.querySelector('.filter-status');
    if (!filterStatusEl) return;
    
    if (filterCount > 0) {
        filterStatusEl.textContent = `已应用${filterCount}个筛选条件，显示${resultCount}/${totalCount}条记录`;
        filterStatusEl.classList.remove('hidden');
    } else {
        filterStatusEl.classList.add('hidden');
    }
}

/**
 * 处理统计文件上传
 * @param {Event} event - 文件上传事件
 */
function handleStatsFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    // 更新上传区域显示
    const uploadArea = document.getElementById('upload-area-stats');
    const fileInfo = uploadArea.querySelector('.file-info');
    const fileName = uploadArea.querySelector('.file-name');
    
    // 显示文件名
    if (fileInfo && fileName) {
        uploadArea.classList.add('file-uploaded');
        fileInfo.classList.remove('hidden');
        fileName.textContent = file.name;
    }
    
    // 显示状态消息
    const statusDiv = document.getElementById('stats-status');
    if (statusDiv) {
        statusDiv.textContent = '文件已上传，正在分析可筛选字段...';
        statusDiv.style.color = 'var(--primary)';
    }
    
    // 读取并处理Excel文件
    readExcelFile(file)
        .then(workbook => {
            // 提取所有工作表数据
            const allData = [];
            workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];
                const sheetData = XLSX.utils.sheet_to_json(worksheet);
                
                // 为每条记录添加工作表标识
                const dataWithSheetInfo = sheetData.map(row => ({
                    ...row,
                    '_原始工作表': sheetName
                }));
                
                allData.push(...dataWithSheetInfo);
            });
            
            // 存储数据
            statsData = allData;
            statsWorkbook = workbook;
            
            // 检测可筛选字段
            const filterFields = detectFilterFields(allData);
            
            // 生成筛选界面
            generateFilterUI(filterFields, 'stats-filters');
            
            if (statusDiv) {
                const fieldsCount = Object.keys(filterFields).length;
                statusDiv.textContent = `检测到${fieldsCount}个可筛选字段，文件已上传，请点击"生成统计报表"按钮进行分析`;
                statusDiv.style.color = 'var(--primary)';
            }
            
            // 显示筛选区域
            const filterContainer = document.getElementById('stats-filters-container');
            if (filterContainer) {
                filterContainer.classList.remove('hidden');
            }
        })
        .catch(error => {
            console.error('处理Excel文件出错:', error);
            if (statusDiv) {
                statusDiv.textContent = `处理Excel文件出错: ${error.message}`;
                statusDiv.style.color = 'var(--danger)';
            }
        });
}

/**
 * 处理对比文件上传
 * @param {Event} event - 文件上传事件
 * @param {string} fileType - 文件类型 ('A' 或 'B')
 */
function handleFileUpload(event, fileType) {
    const file = event.target.files[0];
    if (!file) return;
    
    // 更新上传区域显示
    const uploadAreaId = `upload-area-${fileType}`;
    const uploadArea = document.getElementById(uploadAreaId);
    const fileInfo = uploadArea.querySelector('.file-info');
    const fileName = uploadArea.querySelector('.file-name');
    
    if (fileInfo && fileName) {
        uploadArea.classList.add('file-uploaded');
        fileInfo.classList.remove('hidden');
        fileName.textContent = file.name;
    }
    
    // 读取Excel文件
    readExcelFile(file)
        .then(workbook => {
            // 存储工作簿数据
            if (fileType === 'A') {
                workbookA = workbook;
            } else {
                workbookB = workbook;
            }
            
            // 显示工作表列表
            populateSheetsList(workbook, fileType);
            
            // 检查两个文件是否都已上传
            if (workbookA && workbookB) {
                // 从两个工作簿中提取全部数据用于筛选项检测
                const allDataA = extractAllData(workbookA);
                const allDataB = extractAllData(workbookB);
                const combinedData = [...allDataA, ...allDataB];
                
                // 检测可筛选字段
                const filterFields = detectFilterFields(combinedData);
                
                // 生成筛选界面
                generateFilterUI(filterFields, 'compare-filters');
                
                // 显示筛选区域
                const filterContainer = document.getElementById('compare-filters-container');
                if (filterContainer) {
                    filterContainer.classList.remove('hidden');
                }
            }
        })
        .catch(error => {
            console.error(`处理${fileType}文件出错:`, error);
            displayStatus(`处理${fileType}文件出错: ${error.message}`, 'error');
        });
}

/**
 * 从工作簿中提取所有数据
 * @param {Object} workbook - XLSX工作簿对象
 * @returns {Array} 所有工作表数据合并后的数组
 */
function extractAllData(workbook) {
    if (!workbook || !workbook.SheetNames) return [];
    
    const allData = [];
    workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const sheetData = XLSX.utils.sheet_to_json(worksheet);
        
        // 为每条记录添加工作表标识
        const dataWithSheetInfo = sheetData.map(row => ({
            ...row,
            '_原始工作表': sheetName
        }));
        
        allData.push(...dataWithSheetInfo);
    });
    
    return allData;
}

// 全局变量，存储当前统计模式
let currentStatsMode = 'normal'; // 'normal' 或 'template'
// 全局变量，存储模板配置
let statsTemplateConfig = null;
// 全局变量，存储模板文件工作簿
let statsTemplateWorkbook = null;

/**
 * 初始化统计模式选择器
 */
function initModeSelector() {
    const normalModeBtn = document.getElementById('normalModeBtn');
    const templateModeBtn = document.getElementById('templateModeBtn');
    const templateContainer = document.getElementById('template-config-container');
    
    // 点击普通模式按钮
    if (normalModeBtn) {
        normalModeBtn.addEventListener('click', () => {
            // 更新按钮状态
            normalModeBtn.classList.add('active');
            templateModeBtn.classList.remove('active');
            
            // 隐藏模板配置区域
            if (templateContainer) {
                templateContainer.classList.add('hidden');
            }
            
            // 设置当前模式
            currentStatsMode = 'normal';
        });
    }
    
    // 点击模板模式按钮
    if (templateModeBtn) {
        templateModeBtn.addEventListener('click', () => {
            // 更新按钮状态
            templateModeBtn.classList.add('active');
            normalModeBtn.classList.remove('active');
            
            // 尝试加载模板文件
            loadTemplateFile();
            
            // 显示模板配置区域，但隐藏字段选择部分
            if (templateContainer) {
                templateContainer.classList.remove('hidden');
                
                // 隐藏模板字段配置部分
                const fieldsContainer = document.getElementById('template-fields-container');
                if (fieldsContainer) {
                    fieldsContainer.classList.add('hidden');
                }
                
                // 隐藏模板保存/加载按钮
                const templateActions = templateContainer.querySelector('.template-actions');
                if (templateActions) {
                    templateActions.classList.add('hidden');
                }
            }
            
            // 设置当前模式
            currentStatsMode = 'template';
        });
    }
}

/**
 * 尝试加载模板文件
 */
async function loadTemplateFile() {
    const statusDiv = document.getElementById('stats-status');
    
    try {
        if (statusDiv) {
            statusDiv.textContent = '正在查找统计模板文件...';
            statusDiv.style.color = 'var(--primary)';
        }
        
        // 尝试获取统计文件夹中的模板列表
        const templateList = await fetchTemplateList();
        
        if (!templateList || templateList.length === 0) {
            throw new Error('未找到可用的模板文件');
        }
        
        // 更新状态信息
        if (statusDiv) {
            statusDiv.textContent = `找到 ${templateList.length} 个统计模板`;
            statusDiv.style.color = 'var(--success)';
        }
        
        // 显示模板选择界面
        showTemplateSelector(templateList);
        
        // 如果已有统计数据，自动选择第一个模板并生成统计
        if (statsData && statsData.length > 0 && templateList.length === 1) {
            // 直接使用第一个模板
            await loadSelectedTemplate(templateList[0]);
            
            // 延迟执行以确保状态消息可见
            setTimeout(() => {
                generateStatistics();
            }, 1000);
        }
    } catch (error) {
        console.warn('加载模板文件失败:', error);
        if (statusDiv) {
            statusDiv.textContent = '未找到统计模板文件，将使用手动配置模式';
            statusDiv.style.color = 'var(--warning)';
        }
    }
}

/**
 * 获取统计文件夹中的模板文件列表
 * @returns {Promise<Array>} 模板列表
 */
async function fetchTemplateList() {
    try {
        // 模拟获取模板列表，在实际环境中需要通过服务端API或其他方式获取
        // 尝试检查可能存在的模板文件
        const possibleTemplates = [
            '模板.xlsx', 
            '模板1.xlsx', 
            '模板2.xlsx', 
            '模板3.xlsx',
            '学生统计.xlsx',
            '成绩统计.xlsx',
            '班级统计.xlsx',
            '年级统计.xlsx',
            '教师统计.xlsx',
            '学校统计.xlsx'
        ];
        
        const templates = [];
        const fetchPromises = [];
        
        // 创建所有请求的Promise
        for (const templateName of possibleTemplates) {
            const fetchPromise = fetch(`统计/${templateName}`, { method: 'HEAD' })
                .then(response => {
                    if (response.ok) {
                        templates.push({
                            name: templateName,
                            path: `统计/${templateName}`
                        });
                    }
                })
                .catch(() => {
                    // 请求失败，模板不存在，静默忽略
                });
                
            fetchPromises.push(fetchPromise);
        }
        
        // 等待所有请求完成
        await Promise.all(fetchPromises);
        
        // 按照文件名排序
        templates.sort((a, b) => a.name.localeCompare(b.name));
        
        // 加载完成后输出找到的模板
        console.log(`找到 ${templates.length} 个模板文件:`, templates.map(t => t.name).join(', '));
        
        return templates;
    } catch (error) {
        console.error('获取模板列表失败:', error);
        throw error;
    }
}

/**
 * 显示模板选择界面
 * @param {Array} templateList 模板列表
 */
function showTemplateSelector(templateList) {
    const templateContainer = document.getElementById('template-config-container');
    if (!templateContainer) return;
    
    // 确保模板容器可见
    templateContainer.classList.remove('hidden');
    
    // 清除原有内容，保留只有模板选择器
    templateContainer.innerHTML = '';
    
    // 获取或创建模板选择器
    let templateSelector = document.createElement('div');
    templateSelector.id = 'template-selector';
    templateSelector.className = 'template-selector';
    
    // 创建标题
    const title = document.createElement('h4');
    title.textContent = '选择统计模板';
    title.style.fontSize = '18px';
    title.style.marginBottom = '20px';
    templateSelector.appendChild(title);
    
    // 创建说明文本
    const description = document.createElement('p');
    description.className = 'template-description';
    description.innerHTML = `已找到 <strong>${templateList.length}</strong> 个模板文件，请选择要使用的模板：`;
    templateSelector.appendChild(description);
    
    // 创建模板列表容器
    const templateListEl = document.createElement('div');
    templateListEl.className = 'template-list';
    templateSelector.appendChild(templateListEl);
    
    // 添加到模板配置容器
    templateContainer.appendChild(templateSelector);
    
    // 清空现有内容
    templateListEl.innerHTML = '';
    
    // 添加每个模板选项
    templateList.forEach(template => {
        const templateItem = document.createElement('div');
        templateItem.className = 'template-item';
        
        const templateButton = document.createElement('button');
        templateButton.className = 'template-button';
        templateButton.textContent = template.name;
        
        // 添加图标
        const templateIcon = document.createElement('span');
        templateIcon.className = 'template-icon';
        templateIcon.innerHTML = `
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="feather">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
                <line x1="16" y1="13" x2="8" y2="13"></line>
                <line x1="16" y1="17" x2="8" y2="17"></line>
                <polyline points="10 9 9 9 8 9"></polyline>
            </svg>
        `;
        templateButton.prepend(templateIcon);
        
        // 添加点击事件
        templateButton.addEventListener('click', async () => {
            // 选择该模板
            await loadSelectedTemplate(template);
            
            // 如果已有数据，自动生成统计
            if (statsData && statsData.length > 0) {
                generateStatistics();
            }
        });
        
        templateItem.appendChild(templateButton);
        templateListEl.appendChild(templateItem);
    });
    
    // 添加无模板时的提示
    if (templateList.length === 0) {
        const noTemplateMsg = document.createElement('div');
        noTemplateMsg.className = 'no-template-message';
        noTemplateMsg.innerHTML = `
            <div class="warning-icon">
                <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <circle cx="12" cy="12" r="10"></circle>
                    <line x1="12" y1="8" x2="12" y2="12"></line>
                    <line x1="12" y1="16" x2="12.01" y2="16"></line>
                </svg>
            </div>
            <p>未找到任何模板文件</p>
            <p class="hint">请将Excel模板文件放置在"统计"文件夹中</p>
        `;
        templateSelector.appendChild(noTemplateMsg);
    }
}

/**
 * 加载选定的模板
 * @param {Object} template 模板信息
 * @returns {Promise<void>}
 */
async function loadSelectedTemplate(template) {
    const statusDiv = document.getElementById('stats-status');
    
    try {
        if (statusDiv) {
            statusDiv.textContent = `正在加载模板: ${template.name}...`;
            statusDiv.style.color = 'var(--primary)';
        }
        
        // 加载模板文件
        const response = await fetch(template.path);
        
        if (!response.ok) {
            throw new Error(`无法加载模板文件: ${template.name}`);
        }
        
        const data = await response.arrayBuffer();
        statsTemplateWorkbook = XLSX.read(data, { type: 'array' });
        
        // 存储当前选择的模板信息
        statsTemplateConfig = {
            name: template.name,
            path: template.path
        };
        
        if (statusDiv) {
            statusDiv.textContent = `成功加载模板: ${template.name}，将按此模板生成统计`;
            statusDiv.style.color = 'var(--success)';
        }
        
        // 高亮选中的模板按钮
        document.querySelectorAll('.template-button').forEach(btn => {
            if (btn.textContent === template.name) {
                btn.classList.add('active');
            } else {
                btn.classList.remove('active');
            }
        });
        
        // 隐藏手动配置区域（如果有）
        const fieldsContainer = document.getElementById('template-fields-container');
        if (fieldsContainer) {
            fieldsContainer.classList.add('hidden');
        }
        
        return true;
    } catch (error) {
        console.error('加载模板失败:', error);
        if (statusDiv) {
            statusDiv.textContent = `加载模板失败: ${error.message}`;
            statusDiv.style.color = 'var(--danger)';
        }
        return false;
    }
}

/**
 * 生成模板配置字段
 * @param {Array} data - 数据源
 */
function generateTemplateFields(data) {
    if (!data || data.length === 0) return;
    
    const fieldsContainer = document.getElementById('template-fields-container');
    if (!fieldsContainer) return;
    
    // 清空容器
    fieldsContainer.innerHTML = '';
    
    // 获取第一条数据的所有字段
    const firstRecord = data[0];
    const fields = Object.keys(firstRecord).filter(field => !field.startsWith('_'));
    
    // 创建字段选择框
    fields.forEach(field => {
        const fieldItem = document.createElement('div');
        fieldItem.className = 'template-field-item';
        
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.className = 'template-field-checkbox';
        checkbox.id = `template-field-${field}`;
        checkbox.dataset.field = field;
        checkbox.checked = true; // 默认选中所有字段
        
        const label = document.createElement('label');
        label.className = 'template-field-label';
        label.htmlFor = `template-field-${field}`;
        label.textContent = field;
        
        fieldItem.appendChild(checkbox);
        fieldItem.appendChild(label);
        fieldsContainer.appendChild(fieldItem);
    });
    
    // 如果有保存的模板配置，则应用
    if (statsTemplateConfig) {
        applyTemplateConfig(statsTemplateConfig);
    }
}

/**
 * 应用模板配置
 * @param {Object} config - 模板配置
 */
function applyTemplateConfig(config) {
    if (!config || !config.fields) return;
    
    // 遍历所有字段复选框，设置选中状态
    const checkboxes = document.querySelectorAll('.template-field-checkbox');
    checkboxes.forEach(checkbox => {
        const field = checkbox.dataset.field;
        checkbox.checked = config.fields.includes(field);
    });
}

/**
 * 保存统计模板配置
 */
function saveStatsTemplate() {
    // 收集所有选中的字段
    const selectedFields = [];
    const checkboxes = document.querySelectorAll('.template-field-checkbox:checked');
    
    checkboxes.forEach(checkbox => {
        selectedFields.push(checkbox.dataset.field);
    });
    
    if (selectedFields.length === 0) {
        alert('请至少选择一个字段！');
        return;
    }
    
    // 创建模板配置
    const templateConfig = {
        name: prompt('请输入模板名称', '统计模板') || '统计模板',
        fields: selectedFields,
        createdAt: new Date().toISOString()
    };
    
    // 保存到localStorage
    try {
        const templates = JSON.parse(localStorage.getItem('statsTemplates') || '[]');
        templates.push(templateConfig);
        localStorage.setItem('statsTemplates', JSON.stringify(templates));
        
        // 更新当前模板配置
        statsTemplateConfig = templateConfig;
        
        alert(`模板"${templateConfig.name}"已保存成功！`);
    } catch (error) {
        console.error('保存模板失败:', error);
        alert('保存模板失败: ' + error.message);
    }
}

/**
 * 加载统计模板配置
 */
function loadStatsTemplate() {
    try {
        const templates = JSON.parse(localStorage.getItem('statsTemplates') || '[]');
        
        if (templates.length === 0) {
            alert('没有保存的模板！');
            return;
        }
        
        // 创建模板选择界面
        const templateNames = templates.map(t => t.name);
        const selectedIndex = prompt(`请选择要加载的模板（输入序号1-${templates.length}）：\n${
            templateNames.map((name, i) => `${i + 1}. ${name}`).join('\n')
        }`);
        
        if (!selectedIndex || isNaN(selectedIndex) || selectedIndex < 1 || selectedIndex > templates.length) {
            alert('无效的选择！');
            return;
        }
        
        const selectedTemplate = templates[selectedIndex - 1];
        
        // 应用模板配置
        statsTemplateConfig = selectedTemplate;
        applyTemplateConfig(selectedTemplate);
        
        alert(`模板"${selectedTemplate.name}"已加载成功！`);
    } catch (error) {
        console.error('加载模板失败:', error);
        alert('加载模板失败: ' + error.message);
    }
}

/**
 * 根据模板配置分析统计数据
 * @param {Array} data - 要分析的数据
 * @returns {Object} - 分析结果
 */
function analyzeTemplateData(data) {
    const result = {
        totalRecords: data.length,
        sheets: {},
        custom: {} // 用于存储自定义统计结果
    };
    
    // 分析工作表分布
    data.forEach(row => {
        const sheet = row['_原始工作表'] || '未知';
        if (!result.sheets[sheet]) {
            result.sheets[sheet] = 0;
        }
        result.sheets[sheet]++;
    });
    
    // 如果没有模板配置，则返回基本结果
    if (!statsTemplateConfig || !statsTemplateConfig.fields || statsTemplateConfig.fields.length === 0) {
        return result;
    }
    
    // 根据模板中的字段进行统计
    const fields = statsTemplateConfig.fields;
    
    fields.forEach(field => {
        // 如果字段不存在，则跳过
        if (!data[0] || data[0][field] === undefined) return;
        
        // 创建该字段的统计结果容器
        result.custom[field] = {};
        
        // 统计每个唯一值的出现次数
        data.forEach(row => {
            if (row[field] !== undefined) {
                const value = String(row[field]).trim() || '空值';
                if (!result.custom[field][value]) {
                    result.custom[field][value] = 0;
                }
                result.custom[field][value]++;
            }
        });
        
        // 如果字段值可以被转换为数字，尝试生成数值区间统计
        const numericValues = data
            .map(row => row[field])
            .filter(val => val !== undefined && val !== null && !isNaN(parseFloat(val)))
            .map(val => parseFloat(val));
        
        if (numericValues.length > 0) {
            // 创建成绩区间统计
            result.custom[field + '_分数区间'] = {
                '0-59': 0,
                '60-69': 0,
                '70-79': 0,
                '80-89': 0,
                '90-100': 0,
                '其他': 0
            };
            
            numericValues.forEach(value => {
                if (value >= 0 && value < 60) {
                    result.custom[field + '_分数区间']['0-59']++;
                } else if (value >= 60 && value < 70) {
                    result.custom[field + '_分数区间']['60-69']++;
                } else if (value >= 70 && value < 80) {
                    result.custom[field + '_分数区间']['70-79']++;
                } else if (value >= 80 && value < 90) {
                    result.custom[field + '_分数区间']['80-89']++;
                } else if (value >= 90 && value <= 100) {
                    result.custom[field + '_分数区间']['90-100']++;
                } else {
                    result.custom[field + '_分数区间']['其他']++;
                }
            });
            
            // 如果某个区间没有数据，则删除该区间
            Object.keys(result.custom[field + '_分数区间']).forEach(range => {
                if (result.custom[field + '_分数区间'][range] === 0) {
                    delete result.custom[field + '_分数区间'][range];
                }
            });
            
            // 如果所有区间都没有数据，则删除整个分数区间统计
            if (Object.keys(result.custom[field + '_分数区间']).length === 0) {
                delete result.custom[field + '_分数区间'];
            }
        }
    });
    
    return result;
}

/**
 * 根据模板文件生成统计报表
 * @param {Array} data - 原始数据
 */
async function generateStatsByTemplate(data) {
    const statusDiv = document.getElementById('stats-status');
    const resultsContainer = document.getElementById('stats-results-container');
    
    if (!statsTemplateWorkbook || !data || data.length === 0) {
        if (statusDiv) {
            statusDiv.textContent = '无法生成模板统计，缺少必要数据';
            statusDiv.style.color = 'var(--danger)';
        }
        return;
    }
    
    // 检查是否有已选择的模板
    if (!statsTemplateConfig || !statsTemplateConfig.name) {
        if (statusDiv) {
            statusDiv.textContent = '请先选择一个统计模板';
            statusDiv.style.color = 'var(--warning)';
        }
        return;
    }
    
    try {
        if (statusDiv) {
            statusDiv.textContent = `正在解析模板"${statsTemplateConfig.name}"的结构...`;
        }
        
        // 创建新的工作簿用于输出结果
        const outputWorkbook = XLSX.utils.book_new();
        
        // 处理模板中的每个工作表
        for (let i = 0; i < statsTemplateWorkbook.SheetNames.length; i++) {
            const sheetName = statsTemplateWorkbook.SheetNames[i];
            const templateSheet = statsTemplateWorkbook.Sheets[sheetName];
            
            // 将模板工作表转换为JSON以分析结构
            const templateData = XLSX.utils.sheet_to_json(templateSheet, { header: 'A' });
            
            // 深拷贝模板数据作为输出基础
            const outputData = JSON.parse(JSON.stringify(templateData));
            
            // 查找需要填充的单元格和对应的数据公式
            const cellsToFill = [];
            
            templateData.forEach((row, rowIndex) => {
                Object.keys(row).forEach(col => {
                    const cellValue = row[col];
                    if (typeof cellValue === 'string' && cellValue.startsWith('{{') && cellValue.endsWith('}}')) {
                        // 提取公式内容
                        const formula = cellValue.substring(2, cellValue.length - 2).trim();
                        cellsToFill.push({
                            row: rowIndex,
                            col,
                            formula
                        });
                    }
                });
            });
            
            // 填充数据
            if (cellsToFill.length > 0) {
                if (statusDiv) {
                    statusDiv.textContent = `正在处理"${sheetName}"工作表的统计公式...`;
                }
                
                for (const cell of cellsToFill) {
                    // 执行公式并获取结果
                    const result = executeStatsFormula(cell.formula, data);
                    
                    // 更新输出数据
                    outputData[cell.row][cell.col] = result;
                }
            }
            
            // 将处理后的数据转换为工作表
            const outputSheet = XLSX.utils.json_to_sheet(outputData, { skipHeader: true });
            
            // 复制模板中的单元格格式
            if (templateSheet['!cols']) outputSheet['!cols'] = templateSheet['!cols'];
            if (templateSheet['!rows']) outputSheet['!rows'] = templateSheet['!rows'];
            if (templateSheet['!merges']) outputSheet['!merges'] = templateSheet['!merges'];
            
            // 添加到输出工作簿
            XLSX.utils.book_append_sheet(outputWorkbook, outputSheet, sheetName);
        }
        
        // 获取当前日期时间作为文件名
        const now = new Date();
        const timestamp = now.getFullYear().toString() +
            ('0' + (now.getMonth() + 1)).slice(-2) +
            ('0' + now.getDate()).slice(-2) +
            ('0' + now.getHours()).slice(-2) +
            ('0' + now.getMinutes()).slice(-2);
            
        // 从模板名中获取基本名称（移除.xlsx后缀）
        const templateBaseName = statsTemplateConfig.name.replace(/\.xlsx$/i, '');
        
        // 创建友好的文件名
        let outputFileName;
        if (templateBaseName.includes('统计')) {
            // 如果模板名已经包含"统计"字样，则不再添加"统计报表"
            outputFileName = `${templateBaseName}_${timestamp}.xlsx`;
        } else {
            outputFileName = `${templateBaseName}_统计报表_${timestamp}.xlsx`;
        }
        
        // 导出文件
        XLSX.writeFile(outputWorkbook, outputFileName);
        
        if (statusDiv) {
            statusDiv.textContent = `已按模板"${statsTemplateConfig.name}"生成统计报表并自动下载`;
            statusDiv.style.color = 'var(--success)';
        }
        
        // 在页面上也显示统计结果摘要
        if (resultsContainer) {
            resultsContainer.classList.remove('hidden');
            
            // 创建简单的结果摘要
            const summaryHtml = `
                <div class="stats-summary">
                    <h3>统计报表生成结果</h3>
                    <p>已按照模板"${statsTemplateConfig.name}"生成统计报表。</p>
                    <p>共处理 ${data.length} 条数据记录。</p>
                    <p>报表已自动下载到您的计算机。</p>
                    <p class="template-info">模板路径: ${statsTemplateConfig.path}</p>
                </div>
            `;
            
            // 显示结果摘要
            const tableContainer = document.getElementById('stats-table-container');
            if (tableContainer) {
                tableContainer.innerHTML = summaryHtml;
            }
        }
        
    } catch (error) {
        console.error('按模板生成统计报表失败:', error);
        if (statusDiv) {
            statusDiv.textContent = `生成统计报表出错: ${error.message}`;
            statusDiv.style.color = 'var(--danger)';
        }
    }
}

/**
 * 执行统计公式
 * @param {string} formula - 统计公式
 * @param {Array} data - 原始数据
 * @returns {*} - 公式执行结果
 */
function executeStatsFormula(formula, data) {
    try {
        // 分析公式类型
        if (formula.startsWith('COUNT:')) {
            // 计数公式，例如: COUNT:性别=男
            const condition = formula.substring(6).trim();
            return executeCountFormula(condition, data);
        } else if (formula.startsWith('SUM:')) {
            // 求和公式，例如: SUM:数学成绩,性别=女
            const params = formula.substring(4).trim();
            return executeSumFormula(params, data);
        } else if (formula.startsWith('AVG:')) {
            // 平均值公式，例如: AVG:英语成绩,班级=初一(1)班
            const params = formula.substring(4).trim();
            return executeAvgFormula(params, data);
        } else if (formula.startsWith('MAX:')) {
            // 最大值公式，例如: MAX:语文成绩
            const params = formula.substring(4).trim();
            return executeMaxFormula(params, data);
        } else if (formula.startsWith('MIN:')) {
            // 最小值公式，例如: MIN:物理成绩,班级=高二(3)班
            const params = formula.substring(4).trim();
            return executeMinFormula(params, data);
        } else if (formula === 'TOTAL') {
            // 总记录数
            return data.length;
        } else if (formula.includes('=')) {
            // 简单条件表达式，视为COUNT
            return executeCountFormula(formula, data);
        }
        
        // 未识别的公式
        console.warn('未识别的公式:', formula);
        return `未知公式: ${formula}`;
        
    } catch (error) {
        console.error('执行公式出错:', formula, error);
        return `公式错误: ${error.message}`;
    }
}

/**
 * 执行COUNT统计公式
 * @param {string} condition - 条件表达式，例如 "性别=男"
 * @param {Array} data - 原始数据
 * @returns {number} - 计数结果
 */
function executeCountFormula(condition, data) {
    // 解析条件
    const [field, value] = condition.split('=').map(s => s.trim());
    
    // 计数匹配条件的记录
    return data.filter(row => {
        if (row[field] === undefined) return false;
        const rowValue = String(row[field]).trim();
        return rowValue === value;
    }).length;
}

/**
 * 执行SUM统计公式
 * @param {string} params - 参数，例如 "数学成绩,性别=女"
 * @param {Array} data - 原始数据
 * @returns {number} - 求和结果
 */
function executeSumFormula(params, data) {
    // 解析参数
    const parts = params.split(',');
    const field = parts[0].trim();
    
    // 筛选数据
    let filteredData = data;
    if (parts.length > 1) {
        const condition = parts[1].trim();
        const [condField, condValue] = condition.split('=').map(s => s.trim());
        
        filteredData = data.filter(row => {
            if (row[condField] === undefined) return false;
            const rowValue = String(row[condField]).trim();
            return rowValue === condValue;
        });
    }
    
    // 计算总和
    return filteredData.reduce((sum, row) => {
        const value = row[field];
        if (value === undefined || value === null || isNaN(parseFloat(value))) {
            return sum;
        }
        return sum + parseFloat(value);
    }, 0);
}

/**
 * 执行AVG统计公式
 * @param {string} params - 参数，例如 "英语成绩,班级=初一(1)班"
 * @param {Array} data - 原始数据
 * @returns {number} - 平均值结果
 */
function executeAvgFormula(params, data) {
    // 解析参数
    const parts = params.split(',');
    const field = parts[0].trim();
    
    // 筛选数据
    let filteredData = data;
    if (parts.length > 1) {
        const condition = parts[1].trim();
        const [condField, condValue] = condition.split('=').map(s => s.trim());
        
        filteredData = data.filter(row => {
            if (row[condField] === undefined) return false;
            const rowValue = String(row[condField]).trim();
            return rowValue === condValue;
        });
    }
    
    // 获取所有有效数值
    const validValues = filteredData
        .map(row => row[field])
        .filter(value => value !== undefined && value !== null && !isNaN(parseFloat(value)))
        .map(value => parseFloat(value));
    
    if (validValues.length === 0) {
        return 0;
    }
    
    // 计算平均值
    const sum = validValues.reduce((acc, val) => acc + val, 0);
    return parseFloat((sum / validValues.length).toFixed(2));
}

/**
 * 执行MAX统计公式
 * @param {string} params - 参数，例如 "语文成绩" 或 "语文成绩,班级=初一(1)班"
 * @param {Array} data - 原始数据
 * @returns {number} - 最大值结果
 */
function executeMaxFormula(params, data) {
    // 解析参数
    const parts = params.split(',');
    const field = parts[0].trim();
    
    // 筛选数据
    let filteredData = data;
    if (parts.length > 1) {
        const condition = parts[1].trim();
        const [condField, condValue] = condition.split('=').map(s => s.trim());
        
        filteredData = data.filter(row => {
            if (row[condField] === undefined) return false;
            const rowValue = String(row[condField]).trim();
            return rowValue === condValue;
        });
    }
    
    // 获取所有有效数值
    const validValues = filteredData
        .map(row => row[field])
        .filter(value => value !== undefined && value !== null && !isNaN(parseFloat(value)))
        .map(value => parseFloat(value));
    
    if (validValues.length === 0) {
        return 0;
    }
    
    // 获取最大值
    return Math.max(...validValues);
}

/**
 * 执行MIN统计公式
 * @param {string} params - 参数，例如 "物理成绩" 或 "物理成绩,班级=高二(3)班"
 * @param {Array} data - 原始数据
 * @returns {number} - 最小值结果
 */
function executeMinFormula(params, data) {
    // 解析参数
    const parts = params.split(',');
    const field = parts[0].trim();
    
    // 筛选数据
    let filteredData = data;
    if (parts.length > 1) {
        const condition = parts[1].trim();
        const [condField, condValue] = condition.split('=').map(s => s.trim());
        
        filteredData = data.filter(row => {
            if (row[condField] === undefined) return false;
            const rowValue = String(row[condField]).trim();
            return rowValue === condValue;
        });
    }
    
    // 获取所有有效数值
    const validValues = filteredData
        .map(row => row[field])
        .filter(value => value !== undefined && value !== null && !isNaN(parseFloat(value)))
        .map(value => parseFloat(value));
    
    if (validValues.length === 0) {
        return 0;
    }
    
    // 获取最小值
    return Math.min(...validValues);
}