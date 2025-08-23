
// smart_seat_arrangement/frontend/js/main.js
document.addEventListener('DOMContentLoaded', function() {
    let students = [];
    let seatMap = [];
    let scoreColumns = [];
    let selectedAisles = [];
    let restrictions = [];
    
    // 文件上传处理
    document.getElementById('parseBtn').addEventListener('click', parseExcelFile);
    document.getElementById('excelFile').addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (!file) return;
        
        const fileName = file.name.toLowerCase();
        const isValidType = fileName.endsWith('.xlsx') || fileName.endsWith('.xls');
        
        if (!isValidType) {
            showStatus('请上传.xlsx或.xls格式的文件', true);
            return;
        }
        
        showStatus('文件已选择，点击"解析数据"按钮继续', false);
    });

    // 初始化班级布局预览
    initLayoutPreview();
    
    // 监听行列变化
    document.getElementById('rowsInput').addEventListener('change', updateLayoutPreview);
    document.getElementById('colsInput').addEventListener('change', updateLayoutPreview);
    
    // 特殊设置事件监听
    document.getElementById('addRestrictionBtn').addEventListener('click', addRestriction);
    document.getElementById('addFrontRestrictBtn').addEventListener('click', addFrontRestriction);
    document.getElementById('enableSpecialSettings').addEventListener('change', function() {
        document.getElementById('specialSettingsDetails').open = this.checked;
    });

    function initLayoutPreview() {
        updateLayoutPreview();
        
        // 添加过道按钮事件
        document.getElementById('addAisleBtn').addEventListener('click', function() {
            const rows = parseInt(document.getElementById('rowsInput').value) || 5;
            const cols = parseInt(document.getElementById('colsInput').value) || 6;
            
            // 创建过道选择器
            const aisleSelector = document.createElement('div');
            aisleSelector.className = 'aisle-preview';
            aisleSelector.dataset.row = Math.floor(rows/2);
            aisleSelector.dataset.col = Math.floor(cols/2);
            aisleSelector.addEventListener('click', function() {
                this.classList.toggle('selected');
                updateSelectedAisles();
            });
            
            document.getElementById('aislePreviewContainer').appendChild(aisleSelector);
            updateLayoutPreview();
        });
        
        // 清空过道按钮事件
        document.getElementById('clearAislesBtn').addEventListener('click', function() {
            document.getElementById('aislePreviewContainer').innerHTML = '';
            selectedAisles = [];
            updateLayoutPreview();
        });
    }
    
    function updateLayoutPreview() {
        const rows = parseInt(document.getElementById('rowsInput').value) || 5;
        const cols = parseInt(document.getElementById('colsInput').value) || 6;
        const previewContainer = document.getElementById('layoutPreview');
        
        previewContainer.innerHTML = '';
        previewContainer.style.gridTemplateColumns = `repeat(${cols}, 20px)`;
        
        for (let r = 0; r < rows; r++) {
            for (let c = 0; c < cols; c++) {
                const cell = document.createElement('div');
                cell.className = 'preview-cell';
                cell.dataset.row = r;
                cell.dataset.col = c;
                
                // 检查是否是过道
                const isAisle = selectedAisles.some(a => a.row === r && a.col === c);
                if (isAisle) {
                    cell.classList.add('aisle');
                }
                
                previewContainer.appendChild(cell);
            }
        }
    }
    
    function updateSelectedAisles() {
        selectedAisles = [];
        document.querySelectorAll('.aisle-preview.selected').forEach(el => {
            selectedAisles.push({
                row: parseInt(el.dataset.row),
                col: parseInt(el.dataset.col)
            });
        });
        updateLayoutPreview();
    }

    function parseExcelFile() {
        const fileInput = document.getElementById('excelFile');
        const file = fileInput.files[0];
        
        if (!file) {
            showStatus('请先选择Excel文件', true);
            return;
        }

        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // 获取第一个工作表
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet);
                
                if (jsonData.length === 0) {
                    showStatus('Excel文件中没有数据', true);
                    return;
                }
                
                // 解析学生数据
                students = [];
                scoreColumns = [];
                
                // 增强科目识别算法
                const firstRow = jsonData[0];
                const commonSubjects = ['语文', '数学', '英语', '物理', '化学', '生物', '历史', '地理', '政治'];
                const scoreKeywords = ['成绩', '分数', '分', 'Score', 'Grade'];
                
                // 显示所有列名
                console.log('Excel列名:', Object.keys(firstRow));
                
                for (const key in firstRow) {
                    // 排除非成绩列
                    if (key === '姓名' || key === '学号' || key === '特殊需求') continue;
                    
                    // 检查是否为数字或常见科目名称
                    const isNumber = typeof firstRow[key] === 'number';
                    const isCommonSubject = commonSubjects.some(sub => key.includes(sub));
                    const isScoreColumn = scoreKeywords.some(kw => key.includes(kw));
                    
                    if (isNumber || isCommonSubject || isScoreColumn) {
                        scoreColumns.push(key);
                    }
                }
                
                // 显示识别出的科目
                displayRecognizedSubjects(scoreColumns);
                
                // 处理每个学生
                jsonData.forEach((row, index) => {
                    const student = {
                        id: row['学号'] || `ID${index + 1}`,
                        name: row['姓名'] || `学生${index + 1}`,
                        specialNeeds: row['特殊需求'] || '',
                        total: 0,
                        highScores: {},
                        rawData: row // 保存原始数据
                    };
                    
                    // 计算总分和识别高分科目
                    scoreColumns.forEach(col => {
                        student[col] = row[col] || 0;
                        if (typeof student[col] === 'number') {
                            student.total += student[col];
                        }
                    });
                    
                    // 计算平均分
                    const avgScores = {};
                    scoreColumns.forEach(col => {
                        const sum = jsonData.reduce((acc, cur) => acc + (cur[col] || 0), 0);
                        avgScores[col] = sum / jsonData.length;
                    });
                    
                    // 标记高分科目
                    scoreColumns.forEach(col => {
                        if (student[col] > avgScores[col]) {
                            student.highScores[col] = true;
                        }
                    });
                    
                    students.push(student);
                });
                
                // 显示学生列表
                displayStudentList();
                showStatus(`成功解析 ${students.length} 名学生数据`, false);
                document.getElementById('scoreDistributionInfo').classList.remove('hidden');
                
                // 填充特殊设置下拉框
                populateSpecialSettingsDropdowns();
                
            } catch (error) {
                console.error('解析Excel文件时出错:', error);
                showStatus('解析Excel文件时出错，请检查文件格式', true);
            }
        };
        
        reader.onerror = function() {
            showStatus('读取文件时出错', true);
        };
        
        reader.readAsArrayBuffer(file);
    }

    function displayRecognizedSubjects(subjects) {
        const subjectContainer = document.getElementById('subjectContainer');
        const subjectList = document.getElementById('subjectList');
        
        subjectList.innerHTML = '';
        subjectContainer.classList.remove('hidden');
        
        if (subjects.length === 0) {
            const noSubjects = document.createElement('div');
            noSubjects.className = 'text-sm text-gray-500';
            noSubjects.textContent = '未识别到成绩列';
            subjectList.appendChild(noSubjects);
            return;
        }
        
        subjects.forEach(subject => {
            const badge = document.createElement('span');
            badge.className = 'subject-badge';
            badge.textContent = subject;
            subjectList.appendChild(badge);
        });
        
        // 显示科目统计信息
        const infoDiv = document.createElement('div');
        infoDiv.className = 'text-xs text-gray-500 mt-2';
        infoDiv.textContent = `共识别出 ${subjects.length} 个科目`;
        subjectList.appendChild(infoDiv);
    }

    function showStatus(message, isError) {
        const statusElement = document.getElementById('statusMessage');
        statusElement.textContent = message;
        statusElement.classList.remove('hidden');
        statusElement.classList.remove('status-error');
        statusElement.classList.remove('success');
        
        if (isError) {
            statusElement.classList.add('status-error');
            statusElement.classList.add('error');
        } else {
            statusElement.classList.add('success');
        }
    }

    function displayStudentList() {
        const studentList = document.getElementById('studentList');
        const studentListItems = document.getElementById('studentListItems');
        
        studentListItems.innerHTML = '';
        students.forEach(student => {
            const li = document.createElement('li');
            li.className = 'text-sm text-gray-700';
            
            // 创建学生信息容器
            const studentInfo = document.createElement('div');
            studentInfo.className = 'flex justify-between items-center';
            
            // 学生基本信息
            const basicInfo = document.createElement('div');
            basicInfo.innerHTML = `<span class="font-medium">${student.name}</span> (${student.id})`;
            
            // 学生总分
            const totalScore = document.createElement('div');
            totalScore.className = 'text-blue-600 font-medium';
            totalScore.textContent = `总分: ${student.total}`;
            
            // 高分科目标记
            const highScoreSubjects = Object.keys(student.highScores);
            if (highScoreSubjects.length > 0) {
                const highScoreBadge = document.createElement('span');
                highScoreBadge.className = 'ml-2 px-2 py-1 bg-yellow-100 text-yellow-800 text-xs rounded-full';
                highScoreBadge.textContent = `${highScoreSubjects.length}科高分`;
                basicInfo.appendChild(highScoreBadge);
            }
            
            studentInfo.appendChild(basicInfo);
            studentInfo.appendChild(totalScore);
            li.appendChild(studentInfo);
            
            // 添加特殊需求显示
            if (student.specialNeeds) {
                const needsDiv = document.createElement('div');
                needsDiv.className = 'text-xs text-gray-500 mt-1';
                needsDiv.textContent = `特殊需求: ${student.specialNeeds}`;
                li.appendChild(needsDiv);
            }
            
            // 显示所有数据按钮
            const showDataBtn = document.createElement('button');
            showDataBtn.className = 'text-xs text-blue-500 mt-1';
            showDataBtn.textContent = '查看完整数据';
            showDataBtn.addEventListener('click', () => showStudentData(student));
            li.appendChild(showDataBtn);
            
            studentListItems.appendChild(li);
        });
        
        studentList.classList.remove('hidden');
    }
    
    function showStudentData(student) {
        const modal = document.createElement('div');
        modal.className = 'fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50';
        modal.innerHTML = `
            <div class="bg-white rounded-lg p-6 max-w-2xl w-full max-h-[80vh] overflow-auto">
                <div class="flex justify-between items-center mb-4">
                    <h3 class="text-lg font-bold">${student.name} 的完整数据</h3>
                    <button class="text-gray-500 hover:text-gray-700" onclick="this.parentElement.parentElement.parentElement.remove()">
                        &times;
                    </button>
                </div>
                <div class="grid grid-cols-2 gap-4">
                    ${Object.entries(student.rawData).map(([key, value]) => `
                        <div class="border-b pb-2">
                            <div class="text-sm font-medium text-gray-500">${key}</div>
                            <div class="text-sm">${value || '-'}</div>
                        </div>
                    `).join('')}
                </div>
            </div>
        `;
        document.body.appendChild(modal);
    }

    function populateSpecialSettingsDropdowns() {
        const studentA = document.getElementById('studentA');
        const studentB = document.getElementById('studentB');
        const frontRestrictStudent = document.getElementById('frontRestrictStudent');
        
        // 清空现有选项
        studentA.innerHTML = '<option value="">选择学生A</option>';
        studentB.innerHTML = '<option value="">选择学生B</option>';
        frontRestrictStudent.innerHTML = '<option value="">选择学生</option>';
        
        // 添加学生选项
        students.forEach(student => {
            const option = document.createElement('option');
            option.value = student.id;
            option.textContent = `${student.name} (${student.id})`;
            
            studentA.appendChild(option.cloneNode(true));
            studentB.appendChild(option.cloneNode(true));
            frontRestrictStudent.appendChild(option.cloneNode(true));
        });
    }
    
    function addRestriction() {
        const studentA = document.getElementById('studentA').value;
        const studentB = document.getElementById('studentB').value;
        
        if (!studentA || !studentB) {
            showStatus('请选择两个学生', true);
            return;
        }
        
        if (studentA === studentB) {
            showStatus('不能选择同一个学生', true);
            return;
        }
        
        const studentAName = document.getElementById('studentA').options[document.getElementById('studentA').selectedIndex].text;
        const studentBName = document.getElementById('studentB').options[document.getElementById('studentB').selectedIndex].text;
        
        restrictions.push({
            type: 'pair',
            studentA,
            studentB
        });
        
        updateRestrictionsList();
        showStatus(`已添加限制: ${studentAName} 和 ${studentBName} 不能相邻`, false);
    }
    
    function addFrontRestriction() {
        const studentId = document.getElementById('frontRestrictStudent').value;
        const rows = parseInt(document.getElementById('frontRows').value) || 2;
        
        if (!studentId) {
            showStatus('请选择一个学生', true);
            return;
        }
        
        const studentName = document.getElementById('frontRestrictStudent').options[document.getElementById('frontRestrictStudent').selectedIndex].text;
        
        restrictions.push({
            type: 'front',
            studentId,
            rows
        });
        
        updateRestrictionsList();
        showStatus(`已添加限制: ${studentName} 不能坐在前${rows}排`, false);
    }
    
    function updateRestrictionsList() {
        const restrictionsItems = document.getElementById('restrictionsItems');
        restrictionsItems.innerHTML = '';
        
        restrictions.forEach((restriction, index) => {
            const li = document.createElement('li');
            li.className = 'text-sm flex justify-between items-center';
            
            let text = '';
            if (restriction.type === 'pair') {
                const studentAName = document.getElementById('studentA').options[
                    [...document.getElementById('studentA').options].findIndex(opt => opt.value === restriction.studentA)
                ].text;
                const studentBName = document.getElementById('studentB').options[
                    [...document.getElementById('studentB').options].findIndex(opt => opt.value === restriction.studentB)
                ].text;
                text = `${studentAName} 和 ${studentBName} 不能相邻`;
            } else if (restriction.type === 'front') {
                const studentName = document.getElementById('frontRestrictStudent').options[
                    [...document.getElementById('frontRestrictStudent').options].findIndex(opt => opt.value === restriction.studentId)
                ].text;
                text = `${studentName} 不能坐在前${restriction.rows}排`;
            }
            
            li.innerHTML = `
                <span>${text}</span>
                <button class="text-red-500 text-xs" data-index="${index}">删除</button>
            `;
            
            li.querySelector('button').addEventListener('click', function() {
                restrictions.splice(parseInt(this.dataset.index), 1);
                updateRestrictionsList();
            });
            
            restrictionsItems.appendChild(li);
        });
    }

    // 生成座位表
    document.getElementById('generateBtn').addEventListener('click', generateSeatMap);

    function generateSeatMap() {
        if (students.length === 0) {
            showStatus('请先上传并解析学生数据', true);
            return;
        }

        const rows = parseInt(document.getElementById('rowsInput').value) || 5;
        const cols = parseInt(document.getElementById('colsInput').value) || 6;
        
        // 初始化座位表
        seatMap = Array(rows).fill().map(() => Array(cols).fill(null));
        
        // 设置过道
        selectedAisles.forEach(aisle => {
            if (aisle.row < rows && aisle.col < cols) {
                seatMap[aisle.row][aisle.col] = {
                    type: 'aisle'
                };
            }
        });
        
        // 随机分配学生到座位
        const shuffledStudents = [...students].sort(() => Math.random() - 0.5);
        let studentIndex = 0;
        
        for (let r = 0; r < rows; r++) {
            for (let c = 0; c < cols; c++) {
                // 跳过过道
                if (seatMap[r][c] && seatMap[r][c].type === 'aisle') continue;
                
                if (studentIndex < shuffledStudents.length) {
                    seatMap[r][c] = {
                        type: 'student',
                        student: shuffledStudents[studentIndex++]
                    };
                }
            }
        }
        
        // 检查高分科目分布
        checkHighScoreDistribution();
        
        // 检查特殊限制
        checkRestrictions();
        
        // 渲染座位表
        renderSeatMap();
        showStatus('座位表生成成功', false);
    }
    
    function checkRestrictions() {
        const rows = seatMap.length;
        const cols = seatMap[0].length;
        
        restrictions.forEach(restriction => {
            if (restriction.type === 'pair') {
                // 检查两个学生是否相邻
                for (let r = 0; r < rows; r++) {
                    for (let c = 0; c < cols; c++) {
                        const seat = seatMap[r][c];
                        if (!seat || seat.type !== 'student') continue;
                        
                        if (seat.student.id === restriction.studentA || seat.student.id === restriction.studentB) {
                            // 检查相邻座位
                            const neighbors = [
                                r > 0 ? seatMap[r-1][c] : null, // 上
                                r < rows-1 ? seatMap[r+1][c] : null, // 下
                                c > 0 ? seatMap[r][c-1] : null, // 左
                                c < cols-1 ? seatMap[r][c+1] : null // 右
                            ].filter(s => s && s.type === 'student');
                            
                            // 检查是否有另一个学生在相邻座位
                            const hasRestrictedNeighbor = neighbors.some(neighbor => 
                                neighbor.student.id === restriction.studentA || 
                                neighbor.student.id === restriction.studentB
                            );
                            
                            if (hasRestrictedNeighbor) {
                                // 尝试交换位置
                                for (let nr = 0; nr < rows; nr++) {
                                    for (let nc = 0; nc < cols; nc++) {
                                        const targetSeat = seatMap[nr][nc];
                                        if (targetSeat && targetSeat.type === 'student') {
                                            // 检查目标位置是否满足条件
                                            const targetNeighbors = [
                                                nr > 0 ? seatMap[nr-1][nc] : null,
                                                nr < rows-1 ? seatMap[nr+1][nc] : null,
                                                nc > 0 ? seatMap[nr][nc-1] : null,
                                                nc < cols-1 ? seatMap[nr][nc+1] : null
                                            ].filter(s => s && s.type === 'student');
                                            
                                            const hasRestrictedTargetNeighbor = targetNeighbors.some(neighbor => 
                                                neighbor.student.id === restriction.studentA || 
                                                neighbor.student.id === restriction.studentB
                                            );
                                            
                                            if (!hasRestrictedTargetNeighbor) {
                                                // 交换座位
                                                const temp = seatMap[r][c];
                                                seatMap[r][c] = seatMap[nr][nc];
                                                seatMap[nr][nc] = temp;
                                                return;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            } else if (restriction.type === 'front') {
                // 检查学生是否坐在前排
                for (let r = 0; r < restriction.rows; r++) {
                    for (let c = 0; c < cols; c++) {
                        const seat = seatMap[r][c];
                        if (!seat || seat.type !== 'student') continue;
                        
                        if (seat.student.id === restriction.studentId) {
                            // 尝试交换到后排
                            for (let nr = restriction.rows; nr < rows; nr++) {
                                for (let nc = 0; nc < cols; nc++) {
                                    const targetSeat = seatMap[nr][nc];
                                    if (targetSeat && targetSeat.type === 'student') {
                                        // 检查目标学生是否也有前排限制
                                        const targetHasRestriction = restrictions.some(res => 
                                            res.type === 'front' && 
                                            res.studentId === targetSeat.student.id && 
                                            nr < res.rows
                                        );
                                        
                                        if (!targetHasRestriction) {
                                            // 交换座位
                                            const temp = seatMap[r][c];
                                            seatMap[r][c] = seatMap[nr][nc];
                                            seatMap[nr][nc] = temp;
                                            return;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        });
    }

    function checkHighScoreDistribution() {
        const rows = seatMap.length;
        const cols = seatMap[0].length;
        
        // 检查每个科目
        scoreColumns.forEach(col => {
            // 找出该科目高分学生
            const highScoreStudents = students.filter(s => s.highScores[col]);
            
            // 检查每个高分学生的相邻座位
            for (let r = 0; r < rows; r++) {
                for (let c = 0; c < cols; c++) {
                    const currentSeat = seatMap[r][c];
                    if (!currentSeat || currentSeat.type !== 'student') continue;
                    
                    // 如果是当前科目高分学生
                    if (currentSeat.student.highScores[col]) {
                        // 检查前后左右相邻座位
                        const neighbors = [
                            r > 0 ? seatMap[r-1][c] : null, // 上
                            r < rows-1 ? seatMap[r+1][c] : null, // 下
                            c > 0 ? seatMap[r][c-1] : null, // 左
                            c < cols-1 ? seatMap[r][c+1] : null // 右
                        ].filter(seat => seat && seat.type === 'student');
                        
                        // 统计相邻座位中当前科目高分学生数量
                        const highScoreNeighbors = neighbors.filter(neighbor => 
                            neighbor.student.highScores[col]
                        );
                        
                        // 如果同一科目高分学生相邻超过2个，重新分配
                        if (highScoreNeighbors.length > 2) {
                            // 寻找一个合适的位置交换
                            for (let nr = 0; nr < rows; nr++) {
                                for (let nc = 0; nc < cols; nc++) {
                                    const targetSeat = seatMap[nr][nc];
                                    if (targetSeat && targetSeat.type === 'student') {
                                        // 检查目标位置是否满足条件
                                        const targetNeighbors = [
                                            nr > 0 ? seatMap[nr-1][nc] : null,
                                            nr < rows-1 ? seatMap[nr+1][nc] : null,
                                            nc > 0 ? seatMap[nr][nc-1] : null,
                                            nc < cols-1 ? seatMap[nr][nc+1] : null
                                        ].filter(seat => seat && seat.type === 'student');
                                        
                                        const targetHighScoreNeighbors = targetNeighbors.filter(neighbor => 
                                            neighbor.student.highScores[col]
                                        );
                                        
                                        if (targetHighScoreNeighbors.length <= 1) {
                                            // 交换座位
                                            const temp = seatMap[r][c];
                                            seatMap[r][c] = seatMap[nr][nc];
                                            seatMap[nr][nc] = temp;
                                            return;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        });
    }

    function renderSeatMap() {
        const classroomGrid = document.getElementById('classroomGrid');
        classroomGrid.innerHTML = '';
        
        // 设置网格布局
        const rows = seatMap.length;
        const cols = seatMap[0].length;
        classroomGrid.style.gridTemplateColumns = `repeat(${cols}, minmax(80px, 1fr))`;
        
        // 渲染每个座位
        for (let r = 0; r < rows; r++) {
            for (let c = 0; c < cols; c++) {
                const seat = seatMap[r][c];
                const seatElement = document.createElement('div');
                seatElement.className = 'seat';
                seatElement.dataset.row = r;
                seatElement.dataset.col = c;
                
                if (seat && seat.type === 'student') {
                    seatElement.innerHTML = `
                        <div class="text-center">
                            <div class="font-medium">${seat.student.name}</div>
                            <div class="text-xs text-gray-500">${seat.student.id}</div>
                        </div>
                    `;
                    seatElement.draggable = true;
                    
                    // 拖拽事件
                    seatElement.addEventListener('dragstart', handleDragStart);
                    seatElement.addEventListener('dragover', handleDragOver);
                    seatElement.addEventListener('drop', handleDrop);
                    seatElement.addEventListener('dragend', handleDragEnd);
                } else if (seat && seat.type === 'aisle') {
                    seatElement.className = 'seat aisle';
                    seatElement.innerHTML = '<div class="text-gray-500">过道</div>';
                } else {
                    seatElement.className = 'seat bg-gray-100';
                    seatElement.innerHTML = '<div class="text-gray-400">空座位</div>';
                    seatElement.draggable = true;
                    
                    // 为空白座位添加拖拽事件
                    seatElement.addEventListener('dragstart', handleEmptySeatDragStart);
                    seatElement.addEventListener('dragover', handleDragOver);
                    seatElement.addEventListener('drop', handleDrop);
                    seatElement.addEventListener('dragend', handleDragEnd);
                }
                
                classroomGrid.appendChild(seatElement);
            }
        }
    }

    // 拖拽相关函数
    let draggedSeat = null;
    let isDraggingEmptySeat = false;

    function handleDragStart(e) {
        draggedSeat = this;
        isDraggingEmptySeat = false;
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/html', this.innerHTML);
        this.classList.add('opacity-50', 'border-2', 'border-blue-500');
    }

    function handleEmptySeatDragStart(e) {
        draggedSeat = this;
        isDraggingEmptySeat = true;
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/html', this.innerHTML);
        this.classList.add('opacity-50', 'border-2', 'border-blue-500');
    }

    function handleDragOver(e) {
        if (e.preventDefault) {
            e.preventDefault();
        }
        
        // 检查目标是否是有效座位
        const targetRow = parseInt(this.dataset.row);
        const targetCol = parseInt(this.dataset.col);
        const targetSeat = seatMap[targetRow][targetCol];
        
        // 所有座位都可以放置
        e.dataTransfer.dropEffect = 'move';
        this.classList.add('bg-blue-100');
        
        return false;
    }

    function handleDrop(e) {
        if (e.stopPropagation) {
            e.stopPropagation();
        }
        
        // 获取拖拽源座位
        const sourceRow = parseInt(draggedSeat.dataset.row);
        const sourceCol = parseInt(draggedSeat.dataset.col);
        
        // 获取目标座位
        const targetRow = parseInt(this.dataset.row);
        const targetCol = parseInt(this.dataset.col);
        
        // 交换座位
        if (isDraggingEmptySeat) {
            // 如果是拖动空白座位，相当于将目标座位移动到空白处
            seatMap[sourceRow][sourceCol] = seatMap[targetRow][targetCol];
            seatMap[targetRow][targetCol] = null;
        } else {
            const temp = seatMap[sourceRow][sourceCol];
            seatMap[sourceRow][sourceCol] = seatMap[targetRow][targetCol];
            seatMap[targetRow][targetCol] = temp;
        }
        
        // 重新渲染座位表
        renderSeatMap();
        
        return false;
    }

    function handleDragEnd() {
        this.classList.remove('opacity-50', 'border-2', 'border-blue-500');
        draggedSeat = null;
        isDraggingEmptySeat = false;
    }

    // 导出座位表到Excel
    document.getElementById('exportBtn').addEventListener('click', exportSeatMapToExcel);

    function exportSeatMapToExcel() {
        if (!seatMap || seatMap.length === 0) {
            showStatus('请先生成座位表', true);
            return;
        }

        // 创建工作表数据
        const wsData = [];
        
        // 添加表头行 (列号)
        const headerRow = ['座位表'];
        for (let c = 0; c < seatMap[0].length; c++) {
            headerRow.push(`第${c+1}列`);
        }
        wsData.push(headerRow);

        // 填充座位数据
        for (let r = 0; r < seatMap.length; r++) {
            const rowData = [`第${r+1}排`];
            
            for (let c = 0; c < seatMap[r].length; c++) {
                const seat = seatMap[r][c];
                
                if (seat && seat.type === 'student') {
                    rowData.push(seat.student.name);
                } else if (seat && seat.type === 'aisle') {
                    rowData.push('过道');
                } else {
                    rowData.push('空座位');
                }
            }
            
            wsData.push(rowData);
        }

        // 添加学生详细信息表
        wsData.push([]); // 空行分隔
        wsData.push(['学生详细信息']);
        const detailHeaders = ['姓名', '学号', '总分'];
        scoreColumns.forEach(col => detailHeaders.push(col));
        detailHeaders.push('特殊需求');
        wsData.push(detailHeaders);

        students.forEach(student => {
            const rowData = [
                student.name,
                student.id,
                student.total
            ];
            
            scoreColumns.forEach(col => {
                rowData.push(student[col] || '');
            });
            
            rowData.push(student.specialNeeds || '');
            wsData.push(rowData);
        });

        // 创建工作簿
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        
        // 设置单元格合并 (座位表标题)
        ws['!merges'] = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: seatMap[0].length } },
            { s: { r: seatMap.length + 2, c: 0 }, e: { r: seatMap.length + 2, c: detailHeaders.length - 1 } }
        ];
        
        // 设置列宽
        const colWidths = [];
        for (let i = 0; i <= seatMap[0].length; i++) {
            colWidths.push({ wch: 15 });
        }
        ws['!cols'] = colWidths;

        XLSX.utils.book_append_sheet(wb, ws, "智能座位表");

        // 生成Excel文件并下载
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        saveAs(new Blob([wbout], { type: 'application/octet-stream' }), '智能座位表.xlsx');
        
        showStatus('座位表导出成功', false);
    }

    // 开发中功能提示
    document.querySelectorAll('button').forEach(btn => {
        if (btn.id !== 'parseBtn' && btn.id !== 'exportBtn' && btn.id !== 'generateBtn') {
            btn.addEventListener('click', function() {
                document.getElementById('devAlert').classList.remove('hidden');
                setTimeout(() => {
                    document.getElementById('devAlert').classList.add('hidden');
                }, 3000);
            });
        }
    });
});
