/* globals Chart:false, feather:false, XLSX:false */

(function () {
  'use strict'

  // --- Variáveis de Estado ---
  let myChartInstance = null;
  let lastProcessedChartData = null; 
  
  // Paleta de cores para os gráficos
  const colorPalette = [
      '#0d6efd', '#dc3545', '#198754', '#ffc107', 
      '#6c757d', '#6f42c1', '#20c997', '#0dcaf0', 
      '#641a96', '#00bcd4', '#ff9800', '#8bc34a'
  ];
  
  // Mapeamento para ordenação correta dos meses
  const MONTH_ORDER = {
      'JANEIRO': '01', 'FEVEREIRO': '02', 'MARÇO': '03', 'ABRIL': '04',
      'MAIO': '05', 'JUNHO': '06', 'JULHO': '07', 'AGOSTO': '08',
      'SETEMBRO': '09', 'OUTUBRO': '10', 'NOVEMBRO': '11', 'DEZEMBRO': '12'
  };
  
  // ---------------------------------------------
  // FUNÇÕES AUXILIARES
  // ---------------------------------------------

  /**
   * Tenta extrair o Mês e o Ano do nome do arquivo (ex: "venda (4).xlsx - MARÇO - 2025.csv")
   * Retorna 'MÊS/ANO' ou null.
   */
  function extractMonthYear(fileName) {
    // Expressão regular para capturar MÊS e ANO. 
    // Procura por um dos meses seguido de um hífen e quatro dígitos de ano.
    const match = fileName.match(/(JANEIRO|FEVEREIRO|MARÇO|ABRIL|MAIO|JUNHO|JULHO|AGOSTO|SETEMBRO|OUTUBRO|NOVEMBRO|DEZEMBRO)\s*-\s*(\d{4})/i);
    if (match && match.length >= 3) {
        const month = match[1].toUpperCase();
        const year = match[2];
        return `${month}/${year}`;
    }
    return null;
  }

  /**
   * Cria uma chave ordenável para o mês/ano (ex: 2025-03)
   */
  function getSortableMonthKey(monthYear) {
      if (typeof monthYear !== 'string') return '0000-00';
      const parts = monthYear.split('/');
      if (parts.length === 2) {
          const monthString = parts[0].toUpperCase();
          const year = parts[1];
          const monthNumber = MONTH_ORDER[monthString];
          if (monthNumber) {
              return `${year}-${monthNumber}`;
          }
      }
      return `0000-${monthYear}`; // Fallback para ordenação
  }

  /**
   * Converte a estrutura de dados PIVOT (Múltiplos Cabeçalhos) para uma estrutura PLANA
   * @param {Array<Array>} dataArray Dados da planilha lidos com 'header: 1'
   * @param {string} monthYearKey A chave de Mês/Ano inferida (ex: MARÇO/2025)
   * @returns {Array<Object>} Uma matriz de objetos planos: [{ 'Mês/Ano', 'Subcategoria', 'Quantidade' }]
   */
  function normalizePivotedData(dataArray, monthYearKey) {
      if (!dataArray || dataArray.length < 3) return []; 

      const [headerRow1, headerRow2] = dataArray;
      const dataRows = dataArray.slice(2);
      const normalizedData = [];

      // 1. Mapear Categorias e Subcategorias para cada índice de coluna
      const columnMap = [];
      let currentCategory = '';
      
      // Itera sobre a segunda linha de cabeçalho para obter o nome da Subcategoria (Produto)
      // Começa no índice 1, pois a coluna 0 é o 'Nº SEMANA'
      for (let i = 1; i < headerRow2.length; i++) {
          // Se headerRow1[i] não estiver vazia, é o início de uma nova Categoria (ex: COMPUTADORES)
          if (headerRow1[i] && String(headerRow1[i]).trim().toUpperCase() !== 'Nº SEMANA') {
               currentCategory = String(headerRow1[i]).trim();
          }

          const subcategory = String(headerRow2[i] || '').trim();

          if (subcategory && subcategory.toUpperCase() !== 'Nº SEMANA' && subcategory) {
            // Concatena Categoria e Subcategoria para garantir unicidade, se a categoria não for vazia
            const finalSubcategoryName = currentCategory && currentCategory !== '' 
                                         ? `${currentCategory} - ${subcategory}` 
                                         : subcategory;

            columnMap.push({
                columnIndex: i,
                category: currentCategory,
                subcategory: finalSubcategoryName // Nome completo da subcategoria
            });
          }
      }

      // 2. Iterar sobre as linhas de dados (semanas) e 'desempivotar'
      dataRows.forEach((row) => {
          // Ignorar linhas que parecem ser totais
          if (row[0] && String(row[0]).toUpperCase().includes('TOTAL')) return;
          
          columnMap.forEach(col => {
              // Garante que o valor da célula é tratado como número
              const quantity = parseFloat(row[col.columnIndex]);

              // Somente incluir a linha se houver uma quantidade válida e positiva
              if (!isNaN(quantity) && quantity > 0) {
                  normalizedData.push({
                      'Mês/Ano': monthYearKey,
                      'Subcategoria': col.subcategory, // Usamos apenas a Subcategoria para o gráfico de evolução
                      'Quantidade': quantity,
                  });
              }
          });
      });

      return normalizedData;
  }
  
  /**
   * Consolida todos os dados (já normalizados) por Mês/Ano e Subcategoria.
   */
  function aggregateNormalizedData(data) {
      const groupedData = {}; 
      const totalByCategory = {}; 
      const allSubcategories = new Set();
      const monthlyTotalMap = {}; 

      data.forEach(row => {
          const monthKey = row['Mês/Ano'];
          const subcategory = row['Subcategoria'];
          const quantity = row['Quantidade'];

          if (!monthKey || !subcategory || isNaN(quantity)) return;

          allSubcategories.add(subcategory);
          
          // Agregação por Mês/Ano e Subcategoria
          if (!groupedData[monthKey]) {
              groupedData[monthKey] = {};
          }
          if (!groupedData[monthKey][subcategory]) {
              groupedData[monthKey][subcategory] = 0;
          }
          groupedData[monthKey][subcategory] += quantity;

          // Agregação de Total Geral por Subcategoria
          if (!totalByCategory[subcategory]) {
              totalByCategory[subcategory] = 0;
          }
          totalByCategory[subcategory] += quantity;
          
          // Agregação de Total Mensal para a Tabela
          monthlyTotalMap[monthKey] = (monthlyTotalMap[monthKey] || 0) + quantity;
      });

      // Prepara listas finais ordenadas
      const months = Object.keys(groupedData).sort((a, b) => {
          const keyA = getSortableMonthKey(a);
          const keyB = getSortableMonthKey(b);
          return keyA.localeCompare(keyB);
      });
      const subcategories = Array.from(allSubcategories).sort();

      // Mapeia cores para Subcategorias
      const colorMap = subcategories.reduce((map, sub, index) => {
          map[sub] = colorPalette[index % colorPalette.length];
          return map;
      }, {});

      // Converte o mapa de totais mensais para o formato de tabela
      const monthlyTotals = months.map(month => ({
          'Mês/Ano': month,
          'Quantidade Total': monthlyTotalMap[month].toLocaleString('pt-BR', { maximumFractionDigits: 0 })
      }));

      return { months, subcategories, groupedData, monthlyTotals, totalByCategory, colorMap };
  }
  
  /**
   * Renderiza o gráfico Chart.js
   */
  function renderChart(data, chartType) {
    const chartCanvas = document.getElementById('myChart');
    const initialMessage = document.getElementById('initial-message');
    
    if (!data) return;

    if (myChartInstance) {
        myChartInstance.destroy();
    }
    
    // Esconde a mensagem inicial e exibe o canvas
    if (chartCanvas) {
        chartCanvas.style.display = 'block';
        if (initialMessage) initialMessage.style.display = 'none';
    } else {
        console.error("Elemento Canvas 'myChart' não encontrado.");
        return;
    }

    const datasets = [];
    let config = {};
    let options = {};

    if (chartType === 'bar' || chartType === 'line') {
        // GRÁFICOS DE SÉRIE TEMPORAL (Barra Empilhada e Linha)
        data.subcategories.forEach(sub => {
            const subColor = data.colorMap[sub];
            const subData = data.months.map(month => data.groupedData[month][sub] || 0);

            datasets.push({
                label: sub,
                data: subData,
                backgroundColor: subColor,
                borderColor: subColor,
                fill: chartType === 'line' ? false : true,
                stack: chartType === 'bar' ? 'Stack 1' : undefined,
                type: chartType,
            });
        });

        config = {
            type: 'bar', // Tipo padrão
            data: {
                labels: data.months,
                datasets: datasets,
            },
        };
        
        options = {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                xAxes: [{ 
                    stacked: chartType === 'bar',
                    ticks: { autoSkip: false }
                }],
                yAxes: [{ 
                    stacked: chartType === 'bar',
                    ticks: { beginAtZero: true }
                }]
            },
            tooltips: {
                mode: 'index',
                intersect: false,
            },
            hover: {
                mode: 'nearest',
                intersect: true
            }
        };
    } else {
        // GRÁFICOS DE TOTAL GERAL (Pizza e Rosca)
        const totalLabels = Object.keys(data.totalByCategory);
        const totalQuantities = Object.values(data.totalByCategory);
        const totalColors = totalLabels.map(label => data.colorMap[label]);

        config = {
            type: chartType,
            data: {
                labels: totalLabels,
                datasets: [{
                    data: totalQuantities,
                    backgroundColor: totalColors,
                    borderColor: '#fff',
                    borderWidth: 1,
                }],
            },
        };
        
        options = {
            responsive: true,
            maintainAspectRatio: true, 
            legend: {
                position: 'top',
            },
            title: {
                display: true,
                text: 'Contribuição Total das Subcategorias'
            },
            animation: {
                animateScale: true,
                animateRotate: true
            }
        };
    }

    myChartInstance = new Chart(chartCanvas, { ...config, options: options });
  }

  /**
   * Renderiza a tabela de totais mensais.
   */
  function renderTable(data) {
    const table = document.getElementById('data-table');
    const tableBody = table ? table.querySelector('tbody') : null;

    if (!tableBody || !table) return;
    
    let html = '';
    
    // Recria o cabeçalho para garantir que seja atualizado
    table.innerHTML = '<thead><tr><th scope="col">#</th><th scope="col">Mês/Ano</th><th scope="col">Quantidade Total</th></tr></thead><tbody>';
    
    data.monthlyTotals.forEach((row, index) => {
        html += `
            <tr>
                <td>${index + 1}</td>
                <td>${row['Mês/Ano']}</td>
                <td>${row['Quantidade Total']}</td>
            </tr>
        `;
    });

    if (data.monthlyTotals.length === 0) {
        html = '<tr><td colspan="3">Nenhum dado mensal consolidado encontrado.</td></tr>';
    }

    table.querySelector('tbody').innerHTML = html;
  }
  
  // ---------------------------------------------
  // 1. INICIALIZAÇÃO DO DOM E EVENTOS
  // ---------------------------------------------
  
  window.onload = function() {
      feather.replace({ 'aria-hidden': true } );

      const uploadInput = document.getElementById('excel-upload');
      const chartTypeSelect = document.getElementById('chart-type-select');
      const uploadStatus = document.getElementById('upload-status');
      const errorMessageDiv = document.getElementById('error-message');

      if (!uploadInput || !chartTypeSelect || !uploadStatus) {
        console.error("Um ou mais elementos HTML essenciais não foram encontrados.");
        return;
      }

      // Buffer para acumular dados de múltiplos uploads
      let allNormalizedData = []; 

      
      // Função para processar um único arquivo
      const processFile = (file) => new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = (e) => {
              try {
                  const data = new Uint8Array(e.target.result);
                  // Usar {header: 1} para ler os cabeçalhos em formato de array de arrays (duas linhas de cabeçalho)
                  const workbook = XLSX.read(data, { type: 'array' });
                  
                  const sheetName = workbook.SheetNames[0];
                  const worksheet = workbook.Sheets[sheetName];
                  
                  const dataArray = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                  const monthYearKey = extractMonthYear(file.name);
                  
                  if (!monthYearKey) {
                      // Mensagem de erro aprimorada para guiar o usuário sobre o formato do nome do arquivo
                      if (file.name.toLowerCase().includes('venda.xlsx')) {
                           errorMessageDiv.textContent = `Erro no arquivo "${file.name}": O nome do arquivo é genérico e não contém Mês/Ano. Por favor, renomeie o arquivo para incluir o mês e ano (ex: "vendas - MARÇO - 2025.xlsx").`;
                      } else {
                           errorMessageDiv.textContent = `Erro no arquivo "${file.name}": Não foi possível extrair Mês/Ano do nome do arquivo. O formato esperado é: 'MÊS - AAAA' no nome.`;
                      }
                      
                      // Para a promessa, pois não podemos prosseguir sem a chave de mês/ano
                      reject(); 
                      return;
                  }

                  const normalizedData = normalizePivotedData(dataArray, monthYearKey);
                  
                  if (normalizedData.length === 0) {
                       console.warn(`Arquivo "${file.name}" processado, mas não gerou dados válidos.`);
                       resolve([]); // Resolve com array vazio se não houver dados válidos
                       return;
                  }
                  
                  resolve(normalizedData);
              } catch (error) {
                  errorMessageDiv.textContent = `Erro ao ler ou processar o arquivo "${file.name}". Verifique o formato do XLSX e as células de dados.`;
                  console.error('Erro de processamento XLSX:', error);
                  reject();
              }
          };
          reader.readAsArrayBuffer(file);
      });

      // Manipulador de upload de arquivo
      uploadInput.addEventListener('change', (event) => {
        const files = Array.from(event.target.files);
        if (files.length === 0) return;

        uploadStatus.textContent = `Processando ${files.length} arquivo(s)...`;
        errorMessageDiv.textContent = ''; // Limpa a mensagem de erro anterior
        allNormalizedData = []; // Limpa dados anteriores

        // Processa todos os arquivos em paralelo e espera a conclusão
        Promise.all(files.map(processFile))
            .then(results => {
                // Combina os dados normalizados de todos os arquivos em um único array plano
                allNormalizedData = results.flat().filter(d => d && d['Quantidade'] > 0);
                
                if (allNormalizedData.length > 0) {
                    lastProcessedChartData = aggregateNormalizedData(allNormalizedData);
                    
                    renderChart(lastProcessedChartData, chartTypeSelect.value);
                    renderTable(lastProcessedChartData);
                    
                    uploadStatus.textContent = `Upload e consolidação de ${files.length} arquivo(s) bem-sucedidos!`;
                    errorMessageDiv.textContent = '';
                } else {
                    uploadStatus.textContent = 'Nenhum dado válido encontrado após a consolidação.';
                    // Só exibe a mensagem de erro se a Promise.all() não tiver rejeitado com um erro específico (como nome de arquivo)
                    if (errorMessageDiv.textContent === '') {
                        errorMessageDiv.textContent = 'Verifique se os arquivos estão no formato pivot esperado e contêm valores numéricos.';
                    }
                }
            })
            .catch(() => {
                // O erro específico (como nome de arquivo inválido) já foi setado na função processFile
                uploadStatus.textContent = 'Falha no Upload e processamento de um ou mais arquivos.';
            })
            .finally(() => {
                uploadInput.value = ''; // Limpa a seleção
            });
      });

      // Manipulador de mudança de tipo de gráfico
      chartTypeSelect.addEventListener('change', (event) => {
          if (lastProcessedChartData) {
              renderChart(lastProcessedChartData, event.target.value);
          }
      });
  };
})();