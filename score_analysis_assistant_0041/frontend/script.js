
// score_analysis_assistant_0041/frontend/script.js
document.addEventListener('DOMContentLoaded', function() {
    // 粒子效果初始化
    particlesJS("particles-js", {
        particles: {
            number: { value: 80, density: { enable: true, value_area: 800 } },
            color: { value: "#4facfe" },
            shape: { type: "circle" },
            opacity: { value: 0.5, random: true },
            size: { value: 3, random: true },
            line_linked: { enable: true, distance: 150, color: "#4facfe", opacity: 0.4, width: 1 },
            move: { enable: true, speed: 2, direction: "none", random: true, straight: false, out_mode: "out" }
        },
        interactivity: {
            detect_on: "canvas",
            events: {
                onhover: { enable: true, mode: "grab" },
                onclick: { enable: true, mode: "push" }
            },
            modes: { grab: { distance: 140, line_linked: { opacity: 1 } }, push: { particles_nb: 4 } }
        }
    });

    // 考试类别分数线数据
    const examTypeData = {
        1: { controlLine: 367, privateLine: 389, publicLine: 448 },
        2: { controlLine: 370, privateLine: 398, publicLine: 512 },
        3: { controlLine: 417, privateLine: 440, publicLine: 520 },
        4: { controlLine: 410, privateLine: 430, publicLine: 492 },
        5: { controlLine: 423, privateLine: 433, publicLine: 489 },
        6: { controlLine: 356, privateLine: 371, publicLine: 444 },
        7: { controlLine: 465, privateLine: 480, publicLine: 574 },
        8: { controlLine: 426, privateLine: 441, publicLine: 469 },
        9: { controlLine: 429, privateLine: 440, publicLine: 510 },
        10: { controlLine: 390, privateLine: 407, publicLine: 499 }
    };

    // 模板下载
    document.getElementById('downloadTemplate').addEventListener('click', function() {
        const templateData = [
            ['姓名', '语文', '数学', '专业基础', '职业测试', '总分'],
            ['张三', 120, 135, 85, 280, 620],
            ['李四', 110, 125, 90, 300, 625],
            ['王五', 130, 140, 95, 320, 685],
            ['赵六', 100, 115, 80, 250, 545],
            ['钱七', 90, 105, 75, 230, 500],
            ['孙八', 140, 145, 98, 340, 723],
            ['周九', 85, 95, 70, 210, 460],
            ['吴十', 95, 110, 78, 240, 523]
        ];
        
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(templateData);
        XLSX.utils.book_append_sheet(wb, ws, "成绩模板");
        
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        saveAs(new Blob([wbout], { type: 'application/octet-stream' }), '成绩模板.xlsx');
    });

    // 文件上传处理 - 改为自动触发
    document.getElementById('fileInput').addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (!file) return;
        
        document.getElementById('fileName').textContent = file.name;
        document.getElementById('fileInfo').classList.remove('hidden');
        
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
            
            // 处理数据并显示结果
            processData(jsonData);
            document.getElementById('resultsSection').classList.remove('hidden');
        };
        reader.readAsArrayBuffer(file);
    });

    // 计算统计数据按钮点击事件
    document.getElementById('calculateStats').addEventListener('click', function() {
        const examType = document.getElementById('examType').value;
        const savedData = localStorage.getItem('scoreData');
        
        if (savedData) {
            const data = JSON.parse(savedData);
            calculateExamStats(data, examType);
        } else {
            alert('请先上传成绩文件');
        }
    });

    // 数据处理函数
    function processData(data) {
        // 解析表头
        const headers = data[0];
        const nameIndex = headers.indexOf('姓名');
        const chineseIndex = headers.indexOf('语文');
        const mathIndex = headers.indexOf('数学');
        const basicIndex = headers.indexOf('专业基础');
        const careerIndex = headers.indexOf('职业测试');
        const totalIndex = headers.indexOf('总分');
        
        // 提取数据
        const students = [];
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (row.length > 0) {
                students.push({
                    name: row[nameIndex],
                    chinese: parseFloat(row[chineseIndex]) || 0,
                    math: parseFloat(row[mathIndex]) || 0,
                    basic: parseFloat(row[basicIndex]) || 0,
                    career: parseFloat(row[careerIndex]) || 0,
                    total: parseFloat(row[totalIndex]) || 0
                });
            }
        }
        
        // 生成排名表格
        generateRankTable(students);
        
        // 生成统计数据
        generateStatsTable(students);
        
        // 生成图表
        generateCharts(students);
        
        // 存储数据到localStorage
        localStorage.setItem('scoreData', JSON.stringify(students));
    }
    
    function generateRankTable(data) {
        const subjects = [
            { name: '语文', key: 'chinese' },
            { name: '数学', key: 'math' },
            { name: '专业基础', key: 'basic' },
            { name: '职业测试', key: 'career' },
            { name: '总分', key: 'total' }
        ];
        
        let tableHtml = '';
        
        subjects.forEach(subject => {
            // 排序
            const sorted = [...data].sort((a, b) => b[subject.key] - a[subject.key]);
            
            // 前五名
            const top5 = sorted.slice(0, 5).map(item => `${item.name}(${item[subject.key]})`).join(', ');
            
            // 后五名
            const bottom5 = sorted.slice(-5).map(item => `${item.name}(${item[subject.key]})`).join(', ');
            
            tableHtml += `
                <tr>
                    <td class="py-2 px-4 border-b border-gray-200">${subject.name}</td>
                    <td class="py-2 px-4 border-b border-gray-200">${top5}</td>
                    <td class="py-2 px-4 border-b border-gray-200">${bottom5}</td>
                </tr>
            `;
        });
        
        document.getElementById('rankTableBody').innerHTML = tableHtml;
    }
    
    function generateStatsTable(data) {
        const subjects = [
            { name: '语文', key: 'chinese' },
            { name: '数学', key: 'math' },
            { name: '专业基础', key: 'basic' },
            { name: '职业测试', key: 'career' },
            { name: '总分', key: 'total' }
        ];
        
        let tableHtml = '';
        
        subjects.forEach(subject => {
            const values = data.map(item => item[subject.key]);
            const avg = (values.reduce((a, b) => a + b, 0) / values.length).toFixed(2);
            
            // 计算标准差
            const squareDiffs = values.map(value => Math.pow(value - avg, 2));
            const variance = squareDiffs.reduce((a, b) => a + b, 0) / values.length;
            const stdDev = Math.sqrt(variance).toFixed(2);
            
            tableHtml += `
                <tr>
                    <td class="py-2 px-4 border-b border-gray-200">${subject.name}</td>
                    <td class="py-2 px-4 border-b border-gray-200">${avg}</td>
                    <td class="py-2 px-4 border-b border-gray-200">${stdDev}</td>
                </tr>
            `;
        });
        
        document.getElementById('statsTableBody').innerHTML = tableHtml;
    }
    
    function generateCharts(data) {
        // 语文分数段
        createPieChart('chineseChart', '语文', data.map(item => item.chinese), [
            { label: '80及以下', max: 80 },
            { label: '81-100', min: 81, max: 100 },
            { label: '101-130', min: 101, max: 130 },
            { label: '131及以上', min: 131 }
        ]);
        
        // 数学分数段
        createPieChart('mathChart', '数学', data.map(item => item.math), [
            { label: '80及以下', max: 80 },
            { label: '81-100', min: 81, max: 100 },
            { label: '101-130', min: 101, max: 130 },
            { label: '131及以上', min: 131 }
        ]);
        
        // 专业基础分数段
        createPieChart('basicChart', '专业基础', data.map(item => item.basic), [
            { label: '60及以下', max: 60 },
            { label: '61-80', min: 61, max: 80 },
            { label: '81-90', min: 81, max: 90 },
            { label: '91及以上', min: 91 }
        ]);
        
        // 职业测试分数段
        createPieChart('careerChart', '职业测试', data.map(item => item.career), [
            { label: '180及以下', max: 180 },
            { label: '181-220', min: 181, max: 220 },
            { label: '221-250', min: 221, max: 250 },
            { label: '251-300', min: 251, max: 300 },
            { label: '301及以上', min: 301 }
        ]);
        
        // 总分分数段
        createPieChart('totalChart', '总分', data.map(item => item.total), [
            { label: '280及以下', max: 280 },
            { label: '281-350', min: 281, max: 350 },
            { label: '351-450', min: 351, max: 450 },
            { label: '451-550', min: 451, max: 550 },
            { label: '551-600', min: 551, max: 600 },
            { label: '601-700', min: 601, max: 700 },
            { label: '700及以上', min: 700 }
        ]);
    }
    
    function createPieChart(canvasId, title, scores, ranges) {
        const ctx = document.getElementById(canvasId).getContext('2d');
        
        // 计算各分数段人数
        const counts = ranges.map(range => {
            if (range.min !== undefined && range.max !== undefined) {
                return scores.filter(score => score >= range.min && score <= range.max).length;
            } else if (range.min !== undefined) {
                return scores.filter(score => score >= range.min).length;
            } else {
                return scores.filter(score => score <= range.max).length;
            }
        });
        
        const labels = ranges.map(range => range.label);
        const backgroundColors = [
            'rgba(255, 99, 132, 0.7)',
            'rgba(54, 162, 235, 0.7)',
            'rgba(255, 206, 86, 0.7)',
            'rgba(75, 192, 192, 0.7)',
            'rgba(153, 102, 255, 0.7)',
            'rgba(255, 159, 64, 0.7)',
            'rgba(199, 199, 199, 0.7)'
        ];
        
        new Chart(ctx, {
            type: 'pie',
            data: {
                labels: labels,
                datasets: [{
                    data: counts,
                    backgroundColor: backgroundColors.slice(0, ranges.length),
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: {
                        display: true,
                        text: title + '分数段分布',
                        font: {
                            size: 16
                        }
                    },
                    legend: {
                        position: 'bottom'
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const label = context.label || '';
                                const value = context.raw || 0;
                                const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                const percentage = Math.round((value / total) * 100);
                                return `${label}: ${value}人 (${percentage}%)`;
                            }
                        }
                    }
                }
            }
        });
    }

    // 计算考试统计数据
    function calculateExamStats(data, examType) {
        const totalStudents = data.length;
        const examData = examTypeData[examType];
        
        const passedControl = data.filter(item => item.total >= examData.controlLine).length;
        const passedControlRate = ((passedControl / totalStudents) * 100).toFixed(2) + '%';
        
        const passedPrivate = data.filter(item => item.total >= examData.privateLine).length;
        const passedPrivateRate = ((passedPrivate / totalStudents) * 100).toFixed(2) + '%';
        
        const passedPublic = data.filter(item => item.total >= examData.publicLine).length;
        const passedPublicRate = ((passedPublic / totalStudents) * 100).toFixed(2) + '%';
        
        // 更新表格数据
        document.getElementById('totalStudents').textContent = totalStudents;
        document.getElementById('controlLine').textContent = examData.controlLine;
        document.getElementById('passedControl').textContent = passedControl;
        document.getElementById('passedControlRate').textContent = passedControlRate;
        document.getElementById('privateLine').textContent = examData.privateLine;
        document.getElementById('passedPrivate').textContent = passedPrivate;
        document.getElementById('passedPrivateRate').textContent = passedPrivateRate;
        document.getElementById('publicLine').textContent = examData.publicLine;
        document.getElementById('passedPublic').textContent = passedPublic;
        document.getElementById('passedPublicRate').textContent = passedPublicRate;
    }

    // 从localStorage加载数据
    const savedData = localStorage.getItem('scoreData');
    if (savedData) {
        const data = JSON.parse(savedData);
        document.getElementById('resultsSection').classList.remove('hidden');
        generateRankTable(data);
        generateStatsTable(data);
        generateCharts(data);
    }
});
