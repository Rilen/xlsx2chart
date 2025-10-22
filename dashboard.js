/* globals Chart:false, feather:false, XLSX:false */

(function () {
  'use strict'

  // --- Funções de inicialização ---
  
  // Variável para armazenar a última versão dos dados processados
  let lastProcessedChartData = null; 
  let myChartInstance = null;
  
  const colorPalette = [
      '#0d6efd', '#dc3545', '#198754', '#ffc107', 
      '#6c757d', '#6f42c1', '#20c997', '#0dcaf0', 
      '#641a96', '#00bcd4', '#ff9800', '#8bc34a'
  ];
  
  const MONTH_ORDER = {
      'JANEIRO': '01', 'FEVEREIRO': '02', 'MARÇO': '03', 'ABRIL': '04',
      'MAIO': '05', 'JUNHO': '06', 'JULHO': '07', 'AGOSTO': '08',
      'SETEMBRO': '09', 'OUTUBRO': '10', 'NOVEMBRO': '11', 'DEZEMBRO': '12'
  };
  
  // ---------------------------------------------
  // 1. VARIÁVEIS E INICIALIZAÇÃO DO DOM (GARANTINDO QUE EXISTAM)
  // ---------------------------------------------
  
  // NOVO: Inicializa variáveis DOM dentro de window.onload para garantir que o HTML está carregado
  window.onload = function() {
    
      feather.replace({ 'aria-hidden': true } );

      const uploadInput = document.getElementById('excel-upload');
      const chartTypeSelect = document.getElementById('chart-type-select');
      
      if (!uploadInput) {
          console.error("ERRO FATAL DE INICIALIZAÇÃO: Elemento de upload 'excel-upload' não encontrado no HTML.");
          return;
      }

      uploadInput.addEventListener('change', handleFile, false);

      if (chartTypeSelect) {
          chartTypeSelect.addEventListener('change', () => {
              if (lastProcessedChartData) {
                  // Redesenha usando o novo tipo selecionado
                  drawChart(lastProcessedChartData); 
              }
          });
      }
  };

  // --- Funções principais ---

  function handleFile(e) {
    const files = e.target.files;
    if (files.length === 0) return;
    
    const file = files[0];
    const reader = new FileReader();

    // VARIÁVEIS DOM obtidas aqui dentro da função, onde a execução é garantida
    const statusElement = document.getElementById('upload-status');
    const errorElement = document.getElementById('error-message');
    const table = document.getElementById('data-table');
    const tableBody = table ? table.querySelector('tbody') : null;
    const chartArea = document.getElementById('chart-area');

    statusElement.textContent = `Carregando: ${file.name}...`;
    errorElement.textContent = '';
    if (tableBody) tableBody.innerHTML = '<tr><td colspan="3">Processando...</td></tr>';
    
    if (myChartInstance) myChartInstance.destroy();

    reader.onload = function(event) {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { 
                type: 'array', 
                raw: false, 
                dateNF: 'YYYY-MM-DD' 
            });
            
            const combinedData = [];
            let totalSheets = 0;
            
            workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];
                const sheetData = processSheetData(worksheet, sheetName.trim());
                
                if (sheetData.data.length > 0) {
                    combinedData.push(...sheetData.data);
                    totalSheets++;
                }
            });
            
            if (combinedData.length === 0) {
                 throw new Error(`Nenhum dado válido encontrado nas ${totalSheets} abas processadas.`);
            }
            
            const chartData = prepareDataForStackedBarChart(combinedData);
            lastProcessedChartData = chartData;
            
            // Desenha usando o tipo atual do seletor (Barra Empilhada por padrão)
            drawChart(chartData);
            
            const tableData = aggregateDataForTable(combinedData);
            updateTable(tableData);

            statusElement.textContent = `Sucesso! Carregadas ${totalSheets} meses/categorias (abas) do arquivo "${file.name}".`;

        } catch (error) {
            errorElement.textContent = `ERRO FATAL: ${error.message}`;
            statusElement.textContent = 'Falha no processamento. Verifique o console (F12).';
            if (myChartInstance) myChartInstance.destroy();
            chartArea.innerHTML = '<h3 class="text-center text-muted pt-5">Falha ao processar dados. Verifique a estrutura do Excel.</h3>';
            console.error("Erro no processamento do arquivo Excel:", error);
        }
    };

    reader.readAsArrayBuffer(file);
  }
  
  // ---------------------------------------------
  // FUNÇÃO DE LEITURA DA ABA (ASSUMIDA COMO CORRETA PARA O MODELO DE MULTI-CABEÇALHO)
  // ---------------------------------------------
  function processSheetData(worksheet, monthYear) {
      const dataAsArray = XLSX.utils.sheet_to_json(worksheet, { 
          header: 1, 
          range: 0, 
          raw: false, 
          dateNF: 'YYYY-MM-DD' 
      });
      
      if (dataAsArray.length < 3) return { data: [] };
      
      const subCategoryNames = dataAsArray[1]; 
      const dataStartRowIndex = 2; 
      const periodColIndex = 0; 
      const firstSalesColIndex = 1; 

      const mappedData = [];
      
      for (let i = dataStartRowIndex; i < dataAsArray.length; i++) {
          const row = dataAsArray[i];
          
          if (!row || !row[periodColIndex]) continue;
          
          const weeklyPeriod = row[periodColIndex]; 
          
          for (let j = firstSalesColIndex; j < row.length; j++) {
              const salesValue = parseFloat(row[j]);
              const subcategory = subCategoryNames[j]; 

              if (subcategory && typeof subcategory === 'string' && !subcategory.includes('nan') && !isNaN(salesValue) && salesValue > 0) {
                  mappedData.push({
                      MêsAno: monthYear,
                      Subcategoria: subcategory.trim(),
                      Vendas: salesValue, 
                      FechamentoSemanal: weeklyPeriod 
                  });
              }
          }
      }
      
      return { data: mappedData };
  }

  // ---------------------------------------------
  // 2. AGRUPAMENTO DE DADOS PRINCIPAL (Mantida a lógica de agregação)
  // ---------------------------------------------
  function prepareDataForStackedBarChart(data) {
      const salesByMonthAndSubcategory = {};
      const allMonthsSet = new Set();
      const allSubcategoriesSet = new Set();

      data.forEach(row => {
          const monthKey = row.MêsAno; 
          const subcategory = row.Subcategoria;
          const sales = row.Vendas;

          if (monthKey && subcategory && !isNaN(sales)) {
              allMonthsSet.add(monthKey);
              allSubcategoriesSet.add(subcategory);
              
              const key = `${subcategory}|${monthKey}`;
              salesByMonthAndSubcategory[key] = (salesByMonthAndSubcategory[key] || 0) + sales; 
          }
      });
      
      const sortedMonths = Array.from(allMonthsSet).sort((a, b) => {
          const keyA = getSortableMonthKey(a);
          const keyB = getSortableMonthKey(b);
          return keyA.localeCompare(keyB);
      });
      
      const sortedSubcategories = Array.from(allSubcategoriesSet).sort();
      
      const datasets = sortedSubcategories.map((subcategory, index) => {
          const salesData = sortedMonths.map(month => {
              const key = `${subcategory}|${month}`;
              return salesByMonthAndSubcategory[key] || 0; 
          });
          
          const color = colorPalette[index % colorPalette.length];

          return {
              label: subcategory,
              data: salesData,
              backgroundColor: color,
              borderColor: color,
              borderWidth: 1,
              fill: false 
          };
      });

      const pieData = sortedSubcategories.map(subcategory => {
        let total = 0;
        sortedMonths.forEach(month => {
            const key = `${subcategory}|${month}`;
            total += salesByMonthAndSubcategory[key] || 0;
        });
        return total;
      });
      
      return {
          labels: sortedMonths, 
          datasets: datasets, 
          subcategories: sortedSubcategories, 
          pieData: pieData
      };
  }
  
  // ---------------------------------------------
  // 3. FUNÇÃO PRINCIPAL DE DESENHO
  // ---------------------------------------------
  function drawChart(chartData) {
    const chartTypeSelect = document.getElementById('chart-type-select');
    const chartType = chartTypeSelect ? chartTypeSelect.value : 'bar'; // Padrão é 'bar'
    const chartArea = document.getElementById('chart-area');

    chartArea.innerHTML = '<canvas id="dynamicChart" class="chart-canvas"></canvas>';
    const ctx = document.getElementById('dynamicChart').getContext('2d');
    
    if (myChartInstance) {
      myChartInstance.destroy();
    }
    
    let config;
    
    const baseOptions = {
        responsive: true,
        maintainAspectRatio: false,
        legend: { display: true, position: 'bottom' }
    };
    
    if (['bar', 'line'].includes(chartType)) {
        config = {
            type: chartType,
            data: {
                labels: chartData.labels,
                datasets: chartData.datasets
            },
            options: {
                ...baseOptions,
                title: {
                    display: true,
                    text: 'Evolução Mensal da Quantidade Vendida por Subcategoria'
                },
                scales: {
                    xAxes: [{
                        stacked: chartType === 'bar',
                        scaleLabel: { display: true, labelString: 'Mês / Período' }
                    }],
                    yAxes: [{
                        stacked: chartType === 'bar',
                        ticks: { beginAtZero: true },
                        scaleLabel: { display: true, labelString: 'Quantidade Total Mensal' }
                    }]
                }
            }
        };
    } else if (['pie', 'doughnut'].includes(chartType)) {
        config = {
            type: chartType,
            data: {
                labels: chartData.subcategories,
                datasets: [{
                    data: chartData.pieData,
                    backgroundColor: chartData.subcategories.map((_, i) => colorPalette[i % colorPalette.length]),
                    hoverOffset: 4
                }]
            },
            options: {
                 ...baseOptions,
                 title: {
                     display: true,
                     text: 'Distribuição Total de Quantidade Vendida por Subcategoria'
                 },
                 scales: {}
            }
        };
    }

    if (config) {
        myChartInstance = new Chart(ctx, config);
    }
  }

  // --- Funções Utilitárias ---
  
  function getSortableMonthKey(monthYearString) {
      const parts = monthYearString.toUpperCase().split(/[\s-]+/).filter(p => p);
      if (parts.length >= 2) {
          const monthName = parts[0];
          const year = parts[parts.length - 1];
          const monthCode = MONTH_ORDER[monthName] || '99'; 

          if (year && year.length === 4) {
              return `${year}-${monthCode}`;
          }
      }
      return `0000-${monthYearString}`;
  }

  function aggregateDataForTable(combinedData) {
      const monthlyTotalMap = {};

      combinedData.forEach(row => {
          const monthKey = row.MêsAno;
          const sales = row.Vendas;

          if (monthKey && !isNaN(sales)) {
              monthlyTotalMap[monthKey] = (monthlyTotalMap[monthKey] || 0) + sales;
          }
      });

      return Object.keys(monthlyTotalMap).sort((a, b) => {
          const keyA = getSortableMonthKey(a);
          const keyB = getSortableMonthKey(b);
          return keyA.localeCompare(keyB);
      }).map(month => ({
          MêsAno: month,
          Vendas: monthlyTotalMap[month].toFixed(0)
      }));
  }

  function updateTable(data) {
    const table = document.getElementById('data-table');
    const tableBody = table ? table.querySelector('tbody') : null;

    if (!tableBody || !table) return;
    
    let html = '';
    
    table.innerHTML = '<thead><tr><th scope="col">#</th><th scope="col">Mês/Ano</th><th scope="col">Quantidade Total</th></tr></thead><tbody>';
    
    data.forEach((row, index) => {
        html += `
            <tr>
                <td>${index + 1}</td>
                <td>${row.MêsAno}</td>
                <td>${row.Vendas}</td>
            </tr>
        `;
    });

    if (data.length === 0) {
        html += `<tr><td colspan="3">Nenhum dado mensal consolidado encontrado.</td></tr>`;
    }

    table.querySelector('tbody').innerHTML = html;
  }

})();
