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

    const nomes = {
      card: `REINOW_Card_${clientData.nome.replace(/\s+/g,'_')}.docx`,
      estrategia: `REINOW_Estrategia_${clientData.nome.replace(/\s+/g,'_')}.docx`,
      roteiros: `REINOW_Roteiros_${clientData.nome.replace(/\s+/g,'_')}.docx`,
      legendas: `REINOW_Legendas_${clientData.nome.replace(/\s+/g,'_')}.docx`,
      relatorio: `REINOW_Relatorio_${clientData.nome.replace(/\s+/g,'_')}.docx`,
    };

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${nomes[type]}"`);
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
};

// ── HELPERS ───────────────────────────────────────────────────────────────────
const P = '#4b095d', G = '#d8a43e', LP = '#f0e8f5', LG = '#fdf6e3';
const GR = '#1D9E75', GRL = '#e8f7f2', RD = '#c0392b', RDL = '#fdf0f0';
const W = '#FFFFFF', DK = '#1a1a1a';
const FC = 'Cinzel', FB = 'Josefin Sans';

function nb() { return { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }; }
function noBorders() { return { top: nb(), bottom: nb(), left: nb(), right: nb() }; }
function thinBorder(c='E0D0E8') { return { style: BorderStyle.SINGLE, size: 4, color: c }; }
function thinBorders(c='E0D0E8') { return { top: thinBorder(c), bottom: thinBorder(c), left: thinBorder(c), right: thinBorder(c) }; }
const TW = 10466;
const sp = (b=80,a=80) => ({ spacing: { before: b, after: a } });
const esp = (b=120) => new Paragraph({ children: [new TextRun('')], spacing: { before: b, after: 0 } });

function run(text, opts={}) {
  return new TextRun({ text, font: FB, size: 20, ...opts });
}

function fullRow(text, fill, color, fontName=FC, size=24, bold=true, align=AlignmentType.LEFT) {
  return new Table({ width: { size: TW, type: WidthType.DXA }, columnWidths: [TW], rows: [
    new TableRow({ children: [
      new TableCell({
        borders: noBorders(), width: { size: TW, type: WidthType.DXA },
        shading: { fill: fill.replace('#',''), type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 180, right: 180 },
        children: [new Paragraph({ alignment: align, children: [
          new TextRun({ text, font: fontName, size, bold, color: color.replace('#','') })
        ]})]
      })
    ]})
  ]});
}

function row2(c1fill, c1w, c1items, c2fill, c2items) {
  return new Table({ width: { size: TW, type: WidthType.DXA }, columnWidths: [c1w, TW-c1w], rows: [
    new TableRow({ children: [
      new TableCell({
        borders: noBorders(), width: { size: c1w, type: WidthType.DXA },
        shading: { fill: c1fill.replace('#',''), type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: c1items
      }),
      new TableCell({
        borders: noBorders(), width: { size: TW-c1w, type: WidthType.DXA },
        shading: { fill: c2fill.replace('#',''), type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 160, right: 160 },
        children: c2items
      })
    ]})
  ]});
}

function infoBox(text, fill, borderColor, fontName=FB, size=18) {
  const h = fill.replace('#','');
  const b = borderColor.replace('#','');
  return new Table({ width: { size: TW, type: WidthType.DXA }, columnWidths: [120, TW-120], rows: [
    new TableRow({ children: [
      new TableCell({ borders: noBorders(), width: { size: 120, type: WidthType.DXA }, shading: { fill: b, type: ShadingType.CLEAR }, children: [new Paragraph({ children: [new TextRun('')] })] }),
      new TableCell({ borders: noBorders(), width: { size: TW-120, type: WidthType.DXA }, shading: { fill: h, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 160, right: 160 },
        children: [new Paragraph({ children: [new TextRun({ text, font: fontName, size, color: DK.replace('#',''), italics: true })] })]
      })
    ]})
  ]});
}

function bullet(text, color=G) {
  return new Paragraph({ ...sp(30, 30), children: [
    new TextRun({ text: '● ', font: FC, size: 16, color: color.replace('#','') }),
    new TextRun({ text, font: FB, size: 18, color: DK.replace('#','') })
  ]});
}

function labelPara(text, color=P) {
  return new Paragraph({ ...sp(0, 60), children: [
    new TextRun({ text, font: FC, size: 17, bold: true, color: color.replace('#','') })
  ]});
}

function rodape(nome) {
  return fullRow(`REINOW Marketing  ·  onde o ser humano é rei  ·  Estrategia para ${nome}  ·  @reinowmarketing`, P, G, FC, 16, false, AlignmentType.CENTER);
}

// ── CARD DO CLIENTE ───────────────────────────────────────────────────────────
async function gerarCard(d, extras) {
  const children = [
    // Header
    fullRow(`  REINOW  ·  Card Estratégico de Cliente  ·  ${new Date().toLocaleDateString('pt-BR', {month:'long',year:'numeric'})}`, P, G),
    esp(100),
    // Nome + info
    row2(P, 1800,
      [
        new Paragraph({ ...sp(0,40), children: [new TextRun({ text: d.nome, font: FC, size: 28, bold: true, color: G.replace('#','') })] }),
        new Paragraph({ ...sp(0,40), children: [new TextRun({ text: d.ig ? `@${d.ig}` : '', font: FB, size: 18, color: 'c9a8d4' })] }),
        new Paragraph({ ...sp(0,0), children: [new TextRun({ text: d.nicho || '', font: FB, size: 17, color: 'c9a8d4' })] }),
      ],
      LP,
      [
        labelPara('OBJETIVO NAS REDES'),
        new Paragraph({ ...sp(0,40), children: [new TextRun({ text: d.objetivo || 'Não informado', font: FB, size: 18, color: DK.replace('#','') })] }),
        labelPara('PÚBLICO-ALVO', '#333333'),
        new Paragraph({ ...sp(0,0), children: [new TextRun({ text: extras?.publico || d.raw?.['Quem é a pessoa que mais se beneficia do que você faz?']?.substring(0,120) || 'Ver briefing', font: FB, size: 18, color: DK.replace('#','') })] }),
      ]
    ),
    esp(80),
    // Serviços e ticket
    row2(LG, 4800,
      [
        labelPara('SERVIÇOS / PRODUTOS', P),
        ...(d.raw?.['O que você vende? Liste seus serviços ou produtos principais.'] || 'Ver briefing').split('\n').map(t => bullet(t.trim())).filter((_,i)=>i<5)
      ],
      LP,
      [
        labelPara('TICKET MÉDIO'),
        new Paragraph({ ...sp(0,40), children: [new TextRun({ text: d.raw?.['Qual é o seu ticket médio?'] || 'Não informado', font: FC, size: 26, bold: true, color: P.replace('#','') })] }),
        labelPara('CIDADES'),
        new Paragraph({ ...sp(0,0), children: [new TextRun({ text: d.raw?.['Quais cidades você atua'] || 'Não informado', font: FB, size: 17, color: DK.replace('#','') })] }),
      ]
    ),
    esp(80),
    // História e diferencial
    row2(LP, 5200,
      [
        labelPara('HISTÓRIA E VIRADA'),
        new Paragraph({ ...sp(0,0), children: [new TextRun({ text: (d.raw?.['Por que você escolheu essa área?'] || '').substring(0,300), font: FB, size: 17, color: DK.replace('#','') })] }),
      ],
      LG,
      [
        labelPara('DIFERENCIAL DO MERCADO', P),
        new Paragraph({ ...sp(0,0), children: [new TextRun({ text: (d.raw?.['Cite algo que o seu concorrente ou o mercado faz que a sua empresa não faz'] || '').substring(0,200), font: FB, size: 17, color: DK.replace('#','') })] }),
      ]
    ),
    esp(80),
    // Positivos e negativos
    row2(GRL, 5200,
      [
        labelPara('PONTOS POSITIVOS', GR),
        ...(d.raw?.['Qual foi a primeira vez que fez algo diferente para o seu negócio'] || '').split('\n').slice(0,3).map(t => bullet(t.trim(), GR))
      ],
      RDL,
      [
        labelPara('DESAFIOS / NEGATIVOS', RD),
        ...(d.raw?.['Qual desafio prático você enfrentou e superou'] || '').split('\n').slice(0,3).map(t => bullet(t.trim(), RD))
      ]
    ),
    esp(80),
    // Identidade visual + extras
    row2(P, 3200,
      [
        new Paragraph({ ...sp(0,60), children: [new TextRun({ text: 'IDENTIDADE VISUAL', font: FC, size: 17, bold: true, color: G.replace('#','') })] }),
        new Paragraph({ ...sp(0,30), children: [new TextRun({ text: `Fonte: ${extras?.fonte || 'Não informada'}`, font: FB, size: 17, color: W.replace('#','') })] }),
        new Paragraph({ ...sp(0,0), children: [new TextRun({ text: `Cores: ${extras?.cores || 'Não informadas'}`, font: FB, size: 17, color: W.replace('#','') })] }),
      ],
      LP,
      [
        new Paragraph({ ...sp(0,60), children: [new TextRun({ text: 'INFORMAÇÕES EXTRAS', font: FC, size: 17, bold: true, color: P.replace('#','') })] }),
        new Paragraph({ ...sp(0,30), children: [new TextRun({ text: `ManyChat: ${extras?.manychat || 'Não'}${extras?.palavraChave ? ` · Palavra: ${extras.palavraChave}` : ''}`, font: FB, size: 17, color: DK.replace('#','') })] }),
        new Paragraph({ ...sp(0,0), children: [new TextRun({ text: `Quem posta: ${extras?.quemPosta || 'REINOW'}`, font: FB, size: 17, color: DK.replace('#','') })] }),
        ...(extras?.observacoes ? [new Paragraph({ ...sp(30,0), children: [new TextRun({ text: extras.observacoes, font: FB, size: 16, color: DK.replace('#',''), italics: true })] })] : [])
      ]
    ),
    esp(100),
    rodape(d.nome)
  ];

  const doc = new Document({
    styles: { default: { document: { run: { font: FB, size: 20 } } } },
    sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 720, right: 720, bottom: 720, left: 720 } } }, children }]
  });
  return await Packer.toBuffer(doc);
}

// ── ESTRATÉGIA ────────────────────────────────────────────────────────────────
async function gerarEstrategia(d, texto) {
  const linhas = texto.split('\n').filter(l => l.trim());
  const children = [
    fullRow(`  REINOW Marketing  ·  Estratégia de Conteúdo`, P, G),
    esp(100),
    row2(P, 2000,
      [new Paragraph({ children: [new TextRun({ text: d.nome, font: FC, size: 22, bold: true, color: G.replace('#','') })] })],
      LP,
      [new Paragraph({ children: [new TextRun({ text: d.nicho || '', font: FB, size: 18, color: DK.replace('#','') })] })]
    ),
    esp(120),
    infoBox('Estrategia gerada com base no briefing e nas diretrizes REINOW de conteudo humanizado.', LP, P),
    esp(120),
    ...linhas.map(l => {
      if (l.match(/^#{1,3} /)) return fullRow(`  ${l.replace(/^#+\s*/,'')}`, P, G, FC, 20);
      if (l.startsWith('**') && l.endsWith('**')) return new Paragraph({ ...sp(80,40), children: [new TextRun({ text: l.replace(/\*\*/g,''), font: FC, size: 19, bold: true, color: P.replace('#','') })] });
      if (l.startsWith('- ') || l.startsWith('• ')) return bullet(l.replace(/^[-•]\s*/,''));
      return new Paragraph({ ...sp(40,40), children: [new TextRun({ text: l, font: FB, size: 18, color: DK.replace('#','') })] });
    }),
    esp(160),
    rodape(d.nome)
  ];

  const doc = new Document({
    styles: { default: { document: { run: { font: FB, size: 20 } } } },
    sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 720, right: 720, bottom: 720, left: 720 } } }, children }]
  });
  return await Packer.toBuffer(doc);
}

// ── ROTEIROS ──────────────────────────────────────────────────────────────────
async function gerarRoteiros(d, texto) {
  const linhas = texto.split('\n').filter(l => l.trim());
  const children = [
    fullRow(`  REINOW  ·  Material de Gravação  ·  ${d.nome}`, P, G),
    esp(80),
    row2(LG, 5200,
      [new Paragraph({ children: [new TextRun({ text: 'Como usar este material', font: FC, size: 17, bold: true, color: P.replace('#','') })] }),
       new Paragraph({ ...sp(40,0), children: [new TextRun({ text: 'Leia o gancho em voz alta antes de gravar. Fale como conversa, nao como leitura. Errou? Respira e recomeça.', font: FB, size: 16, color: DK.replace('#','') })] })],
      P,
      [new Paragraph({ children: [new TextRun({ text: 'Tom REINOW', font: FC, size: 17, bold: true, color: G.replace('#','') })] }),
       new Paragraph({ ...sp(40,0), children: [new TextRun({ text: 'Humano · Conversacional · Sem discurso · Como contando para uma amiga', font: FB, size: 16, color: W.replace('#','') })] })]
    ),
    esp(120),
    ...linhas.map(l => {
      if (l.match(/^ROTEIRO\s*\d+/i) || l.match(/^R\d+/)) return fullRow(`  ${l}`, P, G, FC, 19);
      if (l.match(/^(GANCHO|FALA|CTA|CENA|CATEGORIA|FRASE):/i)) {
        const [label, ...rest] = l.split(':');
        return row2(P, 1200,
          [new Paragraph({ children: [new TextRun({ text: label, font: FC, size: 16, bold: true, color: G.replace('#','') })] })],
          'f9f9f9',
          [new Paragraph({ children: [new TextRun({ text: rest.join(':').trim(), font: FB, size: 18, color: DK.replace('#','') })] })]
        );
      }
      if (l.startsWith('- ') || l.startsWith('• ')) return bullet(l.replace(/^[-•]\s*/,''));
      return new Paragraph({ ...sp(30,30), children: [new TextRun({ text: l, font: FB, size: 18, color: DK.replace('#','') })] });
    }),
    esp(160),
    rodape(d.nome)
  ];

  const doc = new Document({
    styles: { default: { document: { run: { font: FB, size: 20 } } } },
    sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 720, right: 720, bottom: 720, left: 720 } } }, children }]
  });
  return await Packer.toBuffer(doc);
}

// ── LEGENDAS + CALENDÁRIO ─────────────────────────────────────────────────────
async function gerarLegendas(d, texto) {
  const linhas = texto.split('\n').filter(l => l.trim());
  const children = [
    fullRow(`  REINOW  ·  Legendas e Calendário  ·  ${d.nome}`, P, G),
    esp(120),
    ...linhas.map(l => {
      if (l.match(/^#{1,3} /)) return fullRow(`  ${l.replace(/^#+\s*/,'')}`, P, G, FC, 20);
      if (l.match(/^\*\*\d{2}\/\d{2}/)) return new Paragraph({ ...sp(100,40), children: [new TextRun({ text: l.replace(/\*\*/g,''), font: FC, size: 18, bold: true, color: P.replace('#','') })] });
      if (l.startsWith('📌') || l.startsWith('📅')) return new Paragraph({ ...sp(60,20), children: [new TextRun({ text: l, font: FB, size: 17, bold: true, color: P.replace('#','') })] });
      return new Paragraph({ ...sp(30,30), children: [new TextRun({ text: l, font: FB, size: 17, color: DK.replace('#','') })] });
    }),
    esp(160),
    rodape(d.nome)
  ];

  const doc = new Document({
    styles: { default: { document: { run: { font: FB, size: 20 } } } },
    sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 720, right: 720, bottom: 720, left: 720 } } }, children }]
  });
  return await Packer.toBuffer(doc);
}

// ── RELATÓRIO ─────────────────────────────────────────────────────────────────
async function gerarRelatorio(d, conteudo) {
  const { metricas, analise, mes } = conteudo;
  const linhas = analise ? analise.split('\n').filter(l => l.trim()) : [];

  const children = [
    fullRow(`  REINOW  ·  Relatório de Resultados  ·  ${mes || ''}`, P, G),
    esp(80),
    row2(P, 2000,
      [new Paragraph({ children: [new TextRun({ text: d.nome, font: FC, size: 22, bold: true, color: G.replace('#','') })] })],
      LP,
      [new Paragraph({ children: [new TextRun({ text: `Período: ${mes || 'Não informado'}`, font: FB, size: 18, color: DK.replace('#','') })] })]
    ),
    esp(100),
    fullRow('  Métricas do Período', P, G, FC, 20),
    esp(60),
    ...Object.entries(metricas || {}).map(([k,v]) =>
      row2(LP, 3000,
        [new Paragraph({ children: [new TextRun({ text: k, font: FC, size: 17, bold: true, color: P.replace('#','') })] })],
        LG,
        [new Paragraph({ children: [new TextRun({ text: String(v), font: FC, size: 22, bold: true, color: P.replace('#','') })] })]
      )
    ),
    esp(100),
    fullRow('  Análise dos Resultados', P, G, FC, 20),
    esp(60),
    ...linhas.map(l => {
      if (l.startsWith('**')) return new Paragraph({ ...sp(80,40), children: [new TextRun({ text: l.replace(/\*\*/g,''), font: FC, size: 18, bold: true, color: P.replace('#','') })] });
      if (l.startsWith('- ') || l.startsWith('• ')) return bullet(l.replace(/^[-•]\s*/,''));
      return new Paragraph({ ...sp(40,40), children: [new TextRun({ text: l, font: FB, size: 18, color: DK.replace('#','') })] });
    }),
    esp(160),
    rodape(d.nome)
  ];

  const doc = new Document({
    styles: { default: { document: { run: { font: FB, size: 20 } } } },
    sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 720, right: 720, bottom: 720, left: 720 } } }, children }]
  });
  return await Packer.toBuffer(doc);
}
