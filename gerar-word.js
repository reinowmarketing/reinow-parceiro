const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType } = require('docx');

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });
  const { type, clientData, extras, conteudo } = req.body;
  try {
    let buffer;
    if (type === 'card') buffer = await gerarCard(clientData, extras);
    else if (type === 'estrategia') buffer = await gerarEstrategia(clientData, conteudo);
    else if (type === 'roteiros') buffer = await gerarRoteiros(clientData, conteudo);
    else if (type === 'legendas') buffer = await gerarLegendas(clientData, conteudo);
    else if (type === 'relatorio') buffer = await gerarRelatorio(clientData, conteudo);
    else return res.status(400).json({ error: 'Tipo invalido' });
    const nome = (clientData.nome || 'cliente').replace(/\s+/g,'_');
    const nomes = { card:`REINOW_Card_${nome}.docx`, estrategia:`REINOW_Estrategia_${nome}.docx`, roteiros:`REINOW_Roteiros_${nome}.docx`, legendas:`REINOW_Legendas_${nome}.docx`, relatorio:`REINOW_Relatorio_${nome}.docx` };
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition',`attachment; filename="${nomes[type]}"`);
    res.send(buffer);
  } catch(err) { console.error(err); res.status(500).json({ error: err.message }); }
};

// ── CORES E HELPERS ───────────────────────────────────────────────────────────
const P='4b095d', G='d8a43e', LP='f0e8f5', LG='fdf6e3';
const GR='1D9E75', GRL='e8f7f2', RD='c0392b', RDL='fdf0f0';
const W='FFFFFF', DK='1a1a1a', MID='888888', PINK='c9a8d4';
const FC='Cinzel', FB='Josefin Sans';
const TW=10466;
const nb=()=>({style:BorderStyle.NONE,size:0,color:'FFFFFF'});
const noBorders=()=>({top:nb(),bottom:nb(),left:nb(),right:nb()});
const thinB=(c='E0D0E8')=>({style:BorderStyle.SINGLE,size:4,color:c});
const thinBorders=(c='E0D0E8')=>({top:thinB(c),bottom:thinB(c),left:thinB(c),right:thinB(c)});
const sp=(b=60,a=60)=>({spacing:{before:b,after:a}});
const esp=(b=100)=>new Paragraph({children:[new TextRun('')],spacing:{before:b,after:0}});

// Busca um campo no raw do briefing por palavras-chave parciais
function getField(raw, keys) {
  if (!raw) return '';
  for (const key of Object.keys(raw)) {
    const kl = key.toLowerCase();
    for (const k of keys) {
      if (kl.includes(k.toLowerCase())) {
        const v = raw[key];
        if (v && v.trim()) return v.trim();
      }
    }
  }
  return '';
}

function fullRow(text, fill, color, fontName=FC, size=22, bold=true, align=AlignmentType.LEFT) {
  return new Table({ width:{size:TW,type:WidthType.DXA}, columnWidths:[TW], rows:[new TableRow({ children:[
    new TableCell({ borders:noBorders(), width:{size:TW,type:WidthType.DXA},
      shading:{fill,type:ShadingType.CLEAR}, margins:{top:120,bottom:120,left:180,right:180},
      children:[new Paragraph({ alignment:align, children:[new TextRun({text,font:fontName,size,bold,color})] })]
    })
  ]})] });
}

function row2(c1fill,c1w,c1ch,c2fill,c2ch,pad=100) {
  const c2w=TW-c1w;
  return new Table({ width:{size:TW,type:WidthType.DXA}, columnWidths:[c1w,c2w], rows:[new TableRow({ children:[
    new TableCell({ borders:noBorders(), width:{size:c1w,type:WidthType.DXA}, shading:{fill:c1fill,type:ShadingType.CLEAR}, margins:{top:pad,bottom:pad,left:140,right:120}, children:c1ch }),
    new TableCell({ borders:noBorders(), width:{size:c2w,type:WidthType.DXA}, shading:{fill:c2fill,type:ShadingType.CLEAR}, margins:{top:pad,bottom:pad,left:160,right:160}, children:c2ch })
  ]})] });
}

function lbl(text, color=P) {
  return new Paragraph({...sp(0,50), children:[new TextRun({text,font:FC,size:17,bold:true,color})]});
}

function txt(text, color=DK, size=18, italic=false) {
  return new Paragraph({...sp(0,40), children:[new TextRun({text:text||'Não informado',font:FB,size,color,italics:italic})]});
}

function bullet(text, dotColor=G, size=17) {
  return new Paragraph({...sp(20,20), children:[
    new TextRun({text:'● ',font:FC,size:15,color:dotColor}),
    new TextRun({text:text||'',font:FB,size,color:DK})
  ]});
}

function secLabel(text) {
  return new Paragraph({...sp(0,60), children:[new TextRun({text:text.toUpperCase(),font:FC,size:17,bold:true,color:G})]});
}

function rodape(nome) {
  return fullRow(`REINOW Marketing  ·  onde o ser humano é rei  ·  ${nome}  ·  @reinowmarketing`, P, G, FC, 16, false, AlignmentType.CENTER);
}

// ── CARD DO CLIENTE ───────────────────────────────────────────────────────────
async function gerarCard(d, extras) {
  const r = d.raw || {};
  const mes = new Date().toLocaleDateString('pt-BR',{month:'long',year:'numeric'});

  // Extrair campos do briefing usando palavras-chave flexíveis
  const servicos   = getField(r,['o que você vende','vende','serviços','produtos','lista seus serviços']);
  const ticket     = getField(r,['ticket médio','ticket medio','valor médio','quanto cobra']);
  const cidades    = getField(r,['quais cidades','cidade','onde atua','cidades']);
  const horarios   = getField(r,['horários','horario','funcionamento','horarios de trabalho']);
  const objetivo   = getField(r,['o que você mais deseja','deseja ter de retorno','objetivo','retorno']);
  const dificul    = getField(r,['maior dificuldade','dificuldade']);
  const apresenta  = getField(r,['como você apresenta','apresenta o que faz','conversa informal']);
  const historia   = getField(r,['por que você escolheu','história','virada','trouxe até aqui']);
  const publico    = getField(r,['quem é a pessoa que mais se beneficia','pessoa','público','descreva ela']);
  const desafio    = getField(r,['desafio prático','desafio','superou']);
  const diferente  = getField(r,['primeira vez que fez algo diferente','algo diferente','deu certo']);
  const facil      = getField(r,['as pessoas de fora acham fácil','acham fácil','complexo']);
  const conselho   = getField(r,['conselho para você mesma','conselho','voltar ao início']);
  const metodo     = getField(r,['método','técnica','prática conhecida','aplica']);
  const acoes      = getField(r,['liste 3 ações','3 ações','tempo real que levam']);
  const rapido     = getField(r,['menos tempo para fazer','menos tempo','imaginam']);
  const vaciladas  = getField(r,['3 vaciladas','vaciladas','cometeu','ensinou']);
  const obvio      = getField(r,['demorou para aprender','parece óbvio','mudança de mentalidade']);
  const discorda   = getField(r,['prática comum','discorda','todo mundo faz','errado']);
  const concorr    = getField(r,['concorrente','mercado faz','empresa não faz']);
  const frases     = getField(r,['complete esta frase','as pessoas acham que','na verdade é sobre']);
  const pergunta   = getField(r,['pergunta','frequência','valor do que você faz']);
  const objeto     = getField(r,['objeto','ferramenta','símbolo','aparece em quase']);
  const bordao     = getField(r,['frase','bordão','expressão','usa muito']);

  // Montar blocos de sessões
  function sessao(titulo, fill, borderColor, itens) {
    const rows = [];
    // Header da sessão
    rows.push(new Table({ width:{size:TW,type:WidthType.DXA}, columnWidths:[TW], rows:[new TableRow({ children:[
      new TableCell({ borders:noBorders(), width:{size:TW,type:WidthType.DXA},
        shading:{fill,type:ShadingType.CLEAR}, margins:{top:80,bottom:80,left:160,right:160},
        children:[new Paragraph({children:[
          new TextRun({text:titulo.toUpperCase(),font:FC,size:17,bold:true,color:borderColor})
        ]})]
      })
    ]})] }));
    // Itens
    itens.filter(([k,v])=>v&&v.trim()).forEach(([k,v])=>{
      rows.push(new Table({ width:{size:TW,type:WidthType.DXA}, columnWidths:[2200,TW-2200], rows:[new TableRow({ children:[
        new TableCell({ borders:thinBorders('F0E0F0'), width:{size:2200,type:WidthType.DXA},
          shading:{fill:'F8F4FC',type:ShadingType.CLEAR}, margins:{top:60,bottom:60,left:120,right:80},
          children:[new Paragraph({children:[new TextRun({text:k,font:FC,size:15,bold:true,color:P})]})]
        }),
        new TableCell({ borders:thinBorders('F0E0F0'), width:{size:TW-2200,type:WidthType.DXA},
          shading:{fill:'FAFAFA',type:ShadingType.CLEAR}, margins:{top:60,bottom:60,left:120,right:120},
          children:[new Paragraph({children:[new TextRun({text:v.substring(0,300),font:FB,size:16,color:DK})]})]
        })
      ]})] }));
    });
    return rows;
  }

  const children = [
    // CABEÇALHO
    new Table({ width:{size:TW,type:WidthType.DXA}, columnWidths:[TW], rows:[new TableRow({ children:[
      new TableCell({ borders:noBorders(), width:{size:TW,type:WidthType.DXA},
        shading:{fill:P,type:ShadingType.CLEAR}, margins:{top:160,bottom:160,left:200,right:200},
        children:[
          new Paragraph({ children:[new TextRun({text:'REINOW  ·  Card Estratégico de Cliente',font:FC,size:22,bold:true,color:G})] }),
          new Paragraph({...sp(40,0), children:[new TextRun({text:`${d.nome}  ·  ${mes}`,font:FB,size:18,color:PINK})] })
        ]
      })
    ]})] }),
    esp(80),

    // BLOCO TOPO: Nome + Objetivo + Público
    row2(P, 3200,
      [
        new Paragraph({...sp(0,40), children:[new TextRun({text:d.nome,font:FC,size:26,bold:true,color:G})]}),
        new Paragraph({...sp(0,30), children:[new TextRun({text:d.ig?`@${d.ig}`:'',font:FB,size:17,color:PINK})]}),
        new Paragraph({...sp(0,0), children:[new TextRun({text:d.nicho||'',font:FB,size:17,color:PINK})]}),
      ],
      LP,
      [
        lbl('Objetivo nas redes'),
        txt(objetivo||d.objetivo),
        lbl('Maior dificuldade'),
        txt(dificul),
      ]
    ),
    esp(60),

    // BLOCO: Serviços + Ticket + Cidades
    row2(LG, 5200,
      [
        lbl('Serviços e produtos', P),
        ...(servicos?servicos.split('\n').filter(l=>l.trim()).slice(0,5).map(l=>bullet(l.trim())):[txt(servicos||'Ver briefing')]),
      ],
      LP,
      [
        lbl('Ticket médio'),
        new Paragraph({...sp(0,40), children:[new TextRun({text:ticket||'Não informado',font:FC,size:24,bold:true,color:P})]}),
        lbl('Cidades de atuação'),
        txt(cidades),
        lbl('Horários'),
        txt(horarios),
      ]
    ),
    esp(60),

    // BLOCO: Identidade visual + Extras
    row2(P, 4200,
      [
        new Paragraph({...sp(0,60), children:[new TextRun({text:'IDENTIDADE VISUAL',font:FC,size:17,bold:true,color:G})]}),
        new Paragraph({...sp(0,30), children:[new TextRun({text:`Fonte: ${extras?.fonte||'Não declarada'}`,font:FB,size:17,color:W})]}),
        new Paragraph({...sp(0,0), children:[new TextRun({text:`Cores: ${extras?.cores||'Não informadas'}`,font:FB,size:17,color:W})]}),
      ],
      LP,
      [
        new Paragraph({...sp(0,60), children:[new TextRun({text:'CONFIGURAÇÕES',font:FC,size:17,bold:true,color:P})]}),
        new Paragraph({...sp(0,30), children:[new TextRun({text:`ManyChat: ${extras?.manychat?`Sim · ${extras.palavraChave}`:'Não'}`,font:FB,size:17,color:DK})]}),
        new Paragraph({...sp(0,30), children:[new TextRun({text:`Quem posta: ${extras?.quemPosta||'REINOW'}`,font:FB,size:17,color:DK})]}),
        ...(extras?.observacoes?[new Paragraph({...sp(30,0), children:[new TextRun({text:extras.observacoes,font:FB,size:16,color:DK,italics:true})]})]:[] ),
      ]
    ),
    esp(80),

    // SESSÕES DO BRIEFING
    fullRow('  BLOCO 1 — Sobre o Negócio', P, G),
    esp(20),
    ...sessao('Dados principais', 'F9F4FC', P, [
      ['Como apresenta','Como apresenta o que faz:'+'\n'+apresenta],
      ['Cidades',cidades],['Horários',horarios],['Ticket médio',ticket],
    ]),
    esp(60),

    fullRow('  BLOCO 2 — História e Identidade', P, G),
    esp(20),
    ...sessao('Origem', 'F9F4FC', P, [
      ['Por que escolheu esta área',historia],
      ['Quem mais se beneficia',publico],
    ]),
    esp(60),

    fullRow('  BLOCO 3 — Desafios e Aprendizados', P, G),
    esp(20),
    ...sessao('Jornada', 'F9F4FC', P, [
      ['Desafio superado',desafio],
      ['Primeira vez que deu certo',diferente],
      ['O que parece fácil mas é complexo',facil],
      ['Conselho para si mesma',conselho],
    ]),
    esp(60),

    fullRow('  BLOCO 4 — Processo e Método', P, G),
    esp(20),
    ...sessao('Operação', 'F9F4FC', P, [
      ['Método/técnica usada',metodo],
      ['3 ações com tempo real',acoes],
      ['O que leva menos tempo',rapido],
    ]),
    esp(60),

    fullRow('  BLOCO 5 — Vaciladas e Posicionamento', P, G),
    esp(20),
    ...sessao('Lições', 'F9F4FC', P, [
      ['Vaciladas e aprendizados',vaciladas],
      ['Algo que demorou aprender',obvio],
      ['Prática que discorda',discorda],
      ['O que o concorrente faz que não faz',concorr],
    ]),
    esp(60),

    fullRow('  BLOCO 6 — Posicionamento e Bordão', P, G),
    esp(20),
    ...sessao('Identidade', 'F9F4FC', P, [
      ['As pessoas acham que é sobre... na verdade é',frases],
      ['Pergunta frequente que mostra falta de percepção de valor',pergunta],
      ['Objeto/símbolo do trabalho',objeto],
      ['Frase/bordão',bordao],
    ]),
    esp(100),

    rodape(d.nome)
  ].flat();

  const doc = new Document({
    styles:{default:{document:{run:{font:FB,size:20}}}},
    sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:720,right:720,bottom:720,left:720}}},children}]
  });
  return await Packer.toBuffer(doc);
}

// ── ESTRATÉGIA ────────────────────────────────────────────────────────────────
async function gerarEstrategia(d, texto) {
  const linhas = (texto||'').split('\n').filter(l=>l.trim());
  const children = [
    fullRow('  REINOW Marketing  ·  Estratégia de Conteúdo', P, G),
    esp(80),
    row2(P, 2200,
      [new Paragraph({children:[new TextRun({text:d.nome,font:FC,size:22,bold:true,color:G})]})],
      LP,
      [new Paragraph({children:[new TextRun({text:d.nicho||'',font:FB,size:18,color:DK})]})],
    ),
    esp(100),
    ...linhas.map(l=>{
      if (l.match(/^#{1,3} /)) return [fullRow(`  ${l.replace(/^#+\s*/,'')}`, P, G, FC, 20), esp(60)];
      if (l.match(/^\*\*(.+)\*\*$/)) return [new Paragraph({...sp(80,40),children:[new TextRun({text:l.replace(/\*\*/g,''),font:FC,size:19,bold:true,color:P})]})];
      if (l.startsWith('- ')||l.startsWith('• ')) return [bullet(l.replace(/^[-•]\s*/,''))];
      return [new Paragraph({...sp(30,30),children:[new TextRun({text:l,font:FB,size:18,color:DK})]})];
    }).flat(),
    esp(140),
    rodape(d.nome)
  ];
  const doc = new Document({styles:{default:{document:{run:{font:FB,size:20}}}},sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:720,right:720,bottom:720,left:720}}},children}]});
  return await Packer.toBuffer(doc);
}

// ── ROTEIROS ──────────────────────────────────────────────────────────────────
async function gerarRoteiros(d, texto) {
  const linhas = (texto||'').split('\n').filter(l=>l.trim());
  const children = [
    fullRow(`  REINOW  ·  Material de Gravação  ·  ${d.nome}`, P, G),
    esp(80),
    row2(LG, 5200,
      [new Paragraph({children:[new TextRun({text:'Como usar',font:FC,size:17,bold:true,color:P})]}),
       new Paragraph({...sp(40,0),children:[new TextRun({text:'Leia o gancho antes de gravar. Fale como conversa. Errou? Respira e recomeça.',font:FB,size:16,color:DK})]})],
      P,
      [new Paragraph({children:[new TextRun({text:'Tom REINOW',font:FC,size:17,bold:true,color:G})]}),
       new Paragraph({...sp(40,0),children:[new TextRun({text:'Humano · Conversacional · Como contando para uma amiga',font:FB,size:16,color:W})]})],
    ),
    esp(100),
    ...linhas.map(l=>{
      if (l.match(/^(ROTEIRO|R)\s*\d+/i)) return [fullRow(`  ${l}`, P, G, FC, 19), esp(40)];
      const m = l.match(/^(GANCHO|FALA|CTA|CENA|CATEGORIA|FRASE|PARTE\s*\d+[^:]*):(.+)/i);
      if (m) return [row2(P, 1200,
        [new Paragraph({children:[new TextRun({text:m[1].trim(),font:FC,size:15,bold:true,color:G})]})],
        'F9F9F9',
        [new Paragraph({children:[new TextRun({text:m[2].trim(),font:FB,size:17,color:DK})]})],
        80
      ), esp(30)];
      if (l.startsWith('- ')||l.startsWith('• ')) return [bullet(l.replace(/^[-•]\s*/,''))];
      return [new Paragraph({...sp(20,20),children:[new TextRun({text:l,font:FB,size:17,color:DK})]})];
    }).flat(),
    esp(140),
    rodape(d.nome)
  ];
  const doc = new Document({styles:{default:{document:{run:{font:FB,size:20}}}},sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:720,right:720,bottom:720,left:720}}},children}]});
  return await Packer.toBuffer(doc);
}

// ── LEGENDAS ──────────────────────────────────────────────────────────────────
async function gerarLegendas(d, texto) {
  const linhas = (texto||'').split('\n').filter(l=>l.trim());
  const children = [
    fullRow(`  REINOW  ·  Legendas e Calendário  ·  ${d.nome}`, P, G),
    esp(100),
    ...linhas.map(l=>{
      if (l.match(/^#{1,3} /)) return [fullRow(`  ${l.replace(/^#+\s*/,'')}`, P, G, FC, 19), esp(50)];
      if (l.match(/^\*\*\d{2}\/\d{2}/)||l.match(/^📅/)) return [new Paragraph({...sp(100,30),children:[new TextRun({text:l.replace(/\*\*/g,''),font:FC,size:18,bold:true,color:P})]})];
      if (l.startsWith('- ')||l.startsWith('• ')) return [bullet(l.replace(/^[-•]\s*/,''))];
      return [new Paragraph({...sp(20,20),children:[new TextRun({text:l,font:FB,size:17,color:DK})]})];
    }).flat(),
    esp(140),
    rodape(d.nome)
  ];
  const doc = new Document({styles:{default:{document:{run:{font:FB,size:20}}}},sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:720,right:720,bottom:720,left:720}}},children}]});
  return await Packer.toBuffer(doc);
}

// ── RELATÓRIO ─────────────────────────────────────────────────────────────────
async function gerarRelatorio(d, conteudo) {
  const { metricas, analise, mes } = conteudo||{};
  const linhas = (analise||'').split('\n').filter(l=>l.trim());
  const children = [
    fullRow(`  REINOW  ·  Relatório de Resultados  ·  ${mes||''}`, P, G),
    esp(80),
    row2(P, 2200,
      [new Paragraph({children:[new TextRun({text:d.nome,font:FC,size:22,bold:true,color:G})]})],
      LP,
      [new Paragraph({children:[new TextRun({text:`Período: ${mes||''}`,font:FB,size:18,color:DK})]})],
    ),
    esp(100),
    fullRow('  Métricas do Período', P, G, FC, 20),
    esp(50),
    ...Object.entries(metricas||{}).filter(([k,v])=>v).map(([k,v])=>
      row2(LP, 3500,
        [new Paragraph({children:[new TextRun({text:k,font:FC,size:17,bold:true,color:P})]})],
        LG,
        [new Paragraph({children:[new TextRun({text:String(v),font:FC,size:20,bold:true,color:P})]})],
      )
    ),
    esp(100),
    fullRow('  Análise dos Resultados', P, G, FC, 20),
    esp(50),
    ...linhas.map(l=>{
      if (l.match(/^#{1,3} /)) return [fullRow(`  ${l.replace(/^#+\s*/,'')}`, P, G, FC, 19), esp(40)];
      if (l.match(/^\*\*(.+)\*\*$/)) return [new Paragraph({...sp(60,30),children:[new TextRun({text:l.replace(/\*\*/g,''),font:FC,size:18,bold:true,color:P})]})];
      if (l.startsWith('- ')||l.startsWith('• ')) return [bullet(l.replace(/^[-•]\s*/,''))];
      return [new Paragraph({...sp(30,30),children:[new TextRun({text:l,font:FB,size:17,color:DK})]})];
    }).flat(),
    esp(140),
    rodape(d.nome)
  ];
  const doc = new Document({styles:{default:{document:{run:{font:FB,size:20}}}},sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:720,right:720,bottom:720,left:720}}},children}]});
  return await Packer.toBuffer(doc);
}
