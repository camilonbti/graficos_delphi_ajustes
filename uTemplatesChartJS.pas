unit uTemplatesChartJS;

interface

uses
  System.SysUtils, System.Classes;

type
  TTemplatesChartJS = class
  public
    class function TemplateBase: string;
    class function TemplateBarras: string;
    class function GetChartJSCode: string;

  end;

implementation

uses uMemo;

class function TTemplatesChartJS.TemplateBase: string;
begin
  Result :=
    '<!DOCTYPE html>' + #13#10 +
    '<html lang="pt-br">' + #13#10 +
    '<head>' + #13#10 +
    '  <meta charset="UTF-8">' + #13#10 +
    '  <meta http-equiv="X-UA-Compatible" content="IE=edge">' + #13#10 +
    '  <meta name="viewport" content="width=device-width, initial-scale=1.0">' + #13#10 +
    '  <title>Chart.js</title>' + #13#10 +
    '  <style>' + #13#10 +
    '    body { margin: 0; padding: 0; }' + #13#10 +
    '    .chart-container { width: 100%; height: 100%; position: relative; }' + #13#10 +
    '  </style>' + #13#10 +
    GetChartJSCode + #13#10 +
    '</head>' + #13#10 +
    '<body>' + #13#10 +
    '  <div class="chart-container">' + #13#10 +
    '    <canvas id="{{GRAFICO_ID}}"></canvas>' + #13#10 +
    '  </div>' + #13#10 +
    '  <script>' + #13#10 +
    '    {{SCRIPT_CONTEUDO}}' + #13#10 +
    '  </script>' + #13#10 +
    '</body>' + #13#10 +
    '</html>';
end;

//class function TTemplatesChartJS.GetChartJSCode: string;
//begin
//  Result := '';
//  Result := Result +'/*!' + sLineBreak;
//  Result := Result +' * Chart.js v3.7.0' + sLineBreak;
//  Result := Result +' * https://www.chartjs.org' + sLineBreak;
//  Result := Result +' * (c) 2022 Chart.js Contributors' + sLineBreak;
//  Result := Result +' * Released under the MIT License' + sLineBreak;
//  Result := Result +' */' + sLineBreak;
//  Result := Result +'!function(t,e){';
//  Result := Result +'"object"==typeof exports&&"undefined"!=typeof module?module.exports=e():';
//  Result := Result +'"function"==typeof define&&define.amd?define(e):';
//  Result := Result +'(t="undefined"!=typeof globalThis?globalThis:t||self).Chart=e()';
//  Result := Result +'}(this,(function(){"use strict";const t="undefined"==typeof window?';
//  Result := Result +'function(t){return t()}:window.requestAnimationFrame;' + sLineBreak;
//
//  // Versão simplificada e condensada do Chart.js
//  Result := Result +'function n(t){return null!==t&&"[object Object]"===Object.prototype.toString.call(t)}' + sLineBreak;
//  Result := Result +'function o(t){return(o="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(t){return typeof t}:function(t){return t&&"function"==typeof Symbol&&t.constructor===Symbol&&t!==Symbol.prototype?"symbol":typeof t})(t)}' + sLineBreak;
//
//  // Implementação básica para Chart
//  Result := Result +'class Chart{constructor(t,e){this.config=e,this.canvas=t.canvas||t,this.ctx=t.getContext?t.getContext("2d"):t.ctx,this.width=t.width||t.clientWidth||';
//                   Result := Result +'t.offsetWidth,this.height=t.height||t.clientHeight||t.offsetHeight,this.id=Math.random().toString(36).substring(2,9),this.data=e.data||{},this.options=e.options||{},this.type=e.type||"bar",this._init()}' + sLineBreak;
//
//  // Inicialização do gráfico
//  Result := Result +'_init(){this._createChart()}' + sLineBreak;
//
//  // Método para criar o gráfico
//  Result := Result +'_createChart(){if(this.type==="bar"){this._drawBar()}else if(this.type==="pie"){this._drawPie()}else{console.warn("Chart type not supported: "+this.type)}}' + sLineBreak;
//
//  // Método para desenhar gráfico de barras
//  Result := Result +'_drawBar(){const ctx=this.ctx;const data=this.data;const options=this.options||{};const chartArea={left:50,top:50,right:this.width-50,';
//  Result := Result +'bottom:this.height-50};const width=chartArea.right-chartArea.left;const height=chartArea.bottom-chartArea.top;const labels=data.labels||[];const datasets=data.datasets||[];' + sLineBreak;
//
//  // Desenho de barras (dividido em partes menores)
//  Result := Result +'const maxValue=Math.max(...datasets.map(d=>Math.max(...d.data)));const barWidth=(width/labels.length)/datasets.length;';
//  Result := Result +'ctx.clearRect(0,0,this.width,this.height);';
//  Result := Result +'datasets.forEach((dataset,datasetIndex)=>{';
//  Result := Result +'dataset.data.forEach((value,index)=>{';
//  Result := Result +'const x=chartArea.left+index*(width/labels.length)+datasetIndex*barWidth;';
//  Result := Result +'const barHeight=(value/maxValue)*height;';
//  Result := Result +'const y=chartArea.bottom-barHeight;';
//  Result := Result +'const color=dataset.backgroundColor?dataset.backgroundColor[index]:this._getDefaultColor(index);';
//  Result := Result +'ctx.fillStyle=color;';
//  Result := Result +'ctx.fillRect(x,y,barWidth,barHeight);';
//  Result := Result +'});});';
//  Result := Result +'}' + sLineBreak;
//
//  // Método para desenhar gráfico de pizza (dividido em partes menores)
//  Result := Result +'_drawPie(){';
//  Result := Result +'const ctx=this.ctx;const data=this.data;const options=this.options||{};';
//  Result := Result +'const centerX=this.width/2;const centerY=this.height/2;';
//  Result := Result +'const radius=Math.min(this.width,this.height)/2-50;';
//  Result := Result +'const labels=data.labels||[];const datasets=data.datasets||[];';
//  Result := Result +'if(!datasets.length)return;';
//  Result := Result +'const values=datasets[0].data;';
//  Result := Result +'const colors=datasets[0].backgroundColor||values.map((_,i)=>this._getDefaultColor(i));';
//  Result := Result +'const total=values.reduce((a,b)=>a+b,0);';
//  Result := Result +'let startAngle=0;';
//  Result := Result +'values.forEach((value,i)=>{';
//  Result := Result +'const sliceAngle=(value/total)*2*Math.PI;';
//  Result := Result +'ctx.beginPath();';
//  Result := Result +'ctx.moveTo(centerX,centerY);';
//  Result := Result +'ctx.arc(centerX,centerY,radius,startAngle,startAngle+sliceAngle);';
//  Result := Result +'ctx.closePath();';
//  Result := Result +'ctx.fillStyle=colors[i];';
//  Result := Result +'ctx.fill();';
//  Result := Result +'startAngle+=sliceAngle;';
//  Result := Result +'});';
//  Result := Result +'}' + sLineBreak;
//
//  // Método para atualizar o gráfico
//  Result := Result +'update(){this.ctx.clearRect(0,0,this.width,this.height);this._createChart()}' + sLineBreak;
//
//  // Método para pegar elementos clicados (dividido em partes menores)
//  Result := Result +'getElementsAtEvent(e){';
//  Result := Result +'const rect=this.canvas.getBoundingClientRect();';
//  Result := Result +'const x=e.clientX-rect.left;';
//  Result := Result +'const y=e.clientY-rect.top;';
//  Result := Result +'if(this.type==="bar"){';
//  Result := Result +'const chartArea={left:50,top:50,right:this.width-50,bottom:this.height-50};';
//  Result := Result +'const width=chartArea.right-chartArea.left;';
//  Result := Result +'const height=chartArea.bottom-chartArea.top;';
//  Result := Result +'const labels=this.data.labels||[];';
//  Result := Result +'if(x>=chartArea.left&&x<=chartArea.right&&y>=chartArea.top&&y<=chartArea.bottom){';
//  Result := Result +'const index=Math.floor((x-chartArea.left)/(width/labels.length));';
//  Result := Result +'return [{_index:index}];';
//  Result := Result +'}}';
//  Result := Result +'else if(this.type==="pie"){';
//  Result := Result +'const centerX=this.width/2;';
//  Result := Result +'const centerY=this.height/2;';
//  Result := Result +'const radius=Math.min(this.width,this.height)/2-50;';
//  Result := Result +'const distanceFromCenter=Math.sqrt(Math.pow(x-centerX,2)+Math.pow(y-centerY,2));';
//  Result := Result +'if(distanceFromCenter<=radius){';
//  Result := Result +'const angle=Math.atan2(y-centerY,x-centerX);';
//  Result := Result +'let startAngle=0;';
//  Result := Result +'const values=this.data.datasets[0].data;';
//  Result := Result +'const total=values.reduce((a,b)=>a+b,0);';
//  Result := Result +'for(let i=0;i<values.length;i++){';
//  Result := Result +'const sliceAngle=(values[i]/total)*2*Math.PI;';
//  Result := Result +'if(angle>=startAngle&&angle<=startAngle+sliceAngle){';
//  Result := Result +'return [{_index:i}];';
//  Result := Result +'}';
//  Result := Result +'startAngle+=sliceAngle;';
//  Result := Result +'}';
//  Result := Result +'}';
//  Result := Result +'}';
//  Result := Result +'return [];';
//  Result := Result +'}' + sLineBreak;
//
//  // Método para gerar cores padrão
//  Result := Result +'_getDefaultColor(index){';
//  Result := Result +'const colors=["rgba(255,99,132,0.6)","rgba(54,162,235,0.6)","rgba(255,206,86,0.6)",';
//  Result := Result +'"rgba(75,192,192,0.6)","rgba(153,102,255,0.6)","rgba(255,159,64,0.6)","rgba(199,199,199,0.6)"];';
//  Result := Result +'return colors[index%colors.length]';
//  Result := Result +'}}' + sLineBreak;
//
//  // Definições estáticas
//  Result := Result +'Chart.defaults={};Chart.instances={};' + sLineBreak;
//
//  // Retorno da função anônima
//  Result := Result +'return Chart;' + sLineBreak;
//
//  // Fechamento da função anônima auto-executável
//  Result := Result +'})());';
//end;

class function TTemplatesChartJS.GetChartJSCode: string;
begin
  Result :=
    '<script>' + #13#10 +
    '  function loadChartJS() {' + #13#10 +
    '    return new Promise((resolve, reject) => {' + #13#10 +
    '      const script = document.createElement("script");' + #13#10 +
    '      script.src = "https://cdn.jsdelivr.net/npm/chart.js@3.7.0/dist/chart.min.js";' + #13#10 +
    '      script.onload = () => {' + #13#10 +
    '        console.log("Chart.js carregado com sucesso");' + #13#10 +
    '        window.chartJSLoaded = true;' + #13#10 +
    '        resolve();' + #13#10 +
    '      };' + #13#10 +
    '      script.onerror = () => {' + #13#10 +
    '        console.error("Erro ao carregar Chart.js");' + #13#10 +
    '        reject(new Error("Erro ao carregar Chart.js"));' + #13#10 +
    '      };' + #13#10 +
    '      document.head.appendChild(script);' + #13#10 +
    '    });' + #13#10 +
    '  }' + #13#10 +
    '</script>' + #13#10;
end;

class function TTemplatesChartJS.TemplateBarras: string;
var
  Base: string;
  Script: string;
begin
  Base := TemplateBase;

  Script :=
    'document.addEventListener("DOMContentLoaded", async function() {' + #13#10 +
    '  try {' + #13#10 +
    '    await loadChartJS();' + #13#10 +
    '    ' + #13#10 +
    '    if (typeof Chart === "undefined") {' + #13#10 +
    '      throw new Error("Chart.js não está disponível mesmo após carregamento");' + #13#10 +
    '    }' + #13#10 +
    '    ' + #13#10 +
    '    var ctx = document.getElementById("{{GRAFICO_ID}}").getContext("2d");' + #13#10 +
    '    if (!ctx) {' + #13#10 +
    '      throw new Error("Canvas não encontrado");' + #13#10 +
    '    }' + #13#10 +
    '    ' + #13#10 +
    '    window.myChart = new Chart(ctx, {' + #13#10 +
    '      type: "bar",' + #13#10 +
    '      data: {{DADOS_JSON}},' + #13#10 +
    '      options: {{CONFIG_JSON}}' + #13#10 +
    '    });' + #13#10 +
    '    ' + #13#10 +
    '    {{SCRIPT_CALLBACK}}' + #13#10 +
    '    ' + #13#10 +
    '    console.log("Gráfico inicializado com sucesso");' + #13#10 +
    '    location.assign("ActionCallBackJS:GraficoCarregado(ok)");' + #13#10 +
    '  } catch(e) {' + #13#10 +
    '    console.error("Erro ao inicializar gráfico:", e);' + #13#10 +
    '    document.body.innerHTML = "<div style=''color:red;padding:20px''>" + e.message + "</div>";' + #13#10 +
    '    location.assign("ActionCallBackJS:Erro(" + e.message + ")");' + #13#10 +
    '  }' + #13#10 +
    '});';

  Result := StringReplace(Base, '{{SCRIPT_CONTEUDO}}', Script, [rfReplaceAll]);
end;

end.




//class function TTemplatesChartJS.GetChartJSCode: string;
//begin
//  Result := '';
//  Result := Result + 'function n(t){return null!==t&&"[object Object]"===Object.prototype.toString.call(t)}' + sLineBreak;
//  Result := Result + 'function o(t){return(o="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(t){return typeof t}:function(t){return t&&"function"==typeof Symbol&&t.constructor===Symbol&&t!==Symbol.prototype?"symbol":typeof t})(t)}' + sLineBreak;
//  Result := Result + 'class Chart{constructor(t,e){this.config=e,this.canvas=t.canvas||t,this.ctx=t.getContext?t.getContext("2d"):t.ctx,this.width=t.width||t.clientWidth||t.offsetWidth,this.height=t.height||t.clientHeight||t.offsetHeight,this.id=Math.random().toString(36).substring(2,9),this.data=e.data||{},this.options=e.options||{},this.type=e.type||"bar",this._init()}' + sLineBreak;
//  Result := Result + '_init(){this._createChart()}' + sLineBreak;
//  Result := Result + '_createChart(){if(this.type==="bar"){this._drawBar()}else if(this.type==="pie"){this._drawPie()}else{console.warn("Chart type not supported: "+this.type)}}' + sLineBreak;
//  Result := Result + '_drawBar(){const ctx=this.ctx;const data=this.data;const options=this.options||{};const chartArea={left:50,top:50,right:this.width-50,bottom:this.height-50};const width=chartArea.right-chartArea.left;const height=chartArea.bottom-chartArea.top;const labels=data.labels||[];const datasets=data.datasets||[];' + sLineBreak;
//end;
