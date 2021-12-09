import * as docx from 'docx';

export function getFigureShiharai(record) {
  const it_kaihatsu_ichiji = String(
    Math.round(record.it_kaihatsu_ichiji.value / 1000)
  ).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
  const it_kaihatsu_uneihi = String(
    Math.round(record.it_kaihatsu_uneihi.value / 1000)
  ).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
  const it_uneihi_ty = String(
    Math.round(record.it_uneihi_ty.value / 1000)
  ).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
  const it_uneihi_ny = String(
    Math.round(record.it_uneihi_ny.value / 1000)
  ).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
  const it_uneihi_n2y = String(
    Math.round(record.it_uneihi_n2y.value / 1000)
  ).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
  const it_uneihi_n3y = String(
    Math.round(record.it_uneihi_n3y.value / 1000)
  ).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
  const hiyo_total = String(Math.round(record.hiyo_total.value / 1000)).replace(
    /(\d)(?=(\d{3})+(?!\d))/g,
    '$1,'
  );
  const kessai_summary = record.kessai_summary.value;
  const project_no = record.project_no.value;
  const project_name = record.project_name.value;
  const system_no = record.system_no.value;
  const haikei = record.haikei.value;
  const jisshi_naiyo = record.jisshi_naiyo.value;
  const hiyo_table = record.hiyo_table.value;
  const file_mitsumori = record.file_mitsumori.value;
  const file_torihikisaki_sentei = record.file_torihikisaki_sentei.value;
  const file_sonota = record.file_sonota.value;

  let kessai_summary_list = kessai_summary.split('\n');
  let paragraph_kessai_summary = [];
  for (let a in kessai_summary_list) {
    paragraph_kessai_summary.push(
      new docx.TextRun({
        text: kessai_summary_list[a],
        break: 1,
      })
    );
  }

  let haikei_list = haikei.split('\n');
  let paragraph_haikei = [];
  for (let a in haikei_list) {
    paragraph_haikei.push(
      new docx.TextRun({
        text: haikei_list[a],
        break: 1,
      })
    );
  }

  let jisshi_naiyo_list = jisshi_naiyo.split('\n');
  let paragraph_jisshi_naiyo = [];
  for (let a in jisshi_naiyo_list) {
    paragraph_jisshi_naiyo.push(
      new docx.TextRun({
        text: jisshi_naiyo_list[a],
        break: 1,
      })
    );
  }

  //
  // 取引テーブルの作成
  //

  //
  // 取引テーブルのヘッダ作成
  //
  function table_torihiki_cell(c_size, c_text) {
    const table_cell = new docx.TableCell({
      width: {
        size: c_size,
        type: docx.WidthType.PERCENTAGE,
      },
      shading: {
        fill: 'F5F5F5',
      },
      children: [
        new docx.Paragraph({
          text: c_text,
          alignment: docx.AlignmentType.CENTER,
        }),
      ],
    });
    return table_cell;
  }

  const table_torihiki = new docx.Table({
    rows: [
      new docx.TableRow({
        children: [
          table_torihiki_cell(40, '取引先'),
          table_torihiki_cell(60, '納入予定日'),
        ],
      }),
    ],
  });

  //
  // 取引テーブルの明細作成
  //
  const groupBy = (array) => {
    return array.reduce((result, currentValue) => {
      result[currentValue.value.hiyo_torihikisaki.value] =
        result[currentValue.value.hiyo_torihikisaki.value] || [];
      var check = result[currentValue.value.hiyo_torihikisaki.value].find(
        (v) => v.nounyuyotei === currentValue.value.hiyo_nounyuyotei.value
      );
      if (!check) {
        result[currentValue.value.hiyo_torihikisaki.value].push({
          nounyuyotei: currentValue.value.hiyo_nounyuyotei.value,
        });
      }
      return result;
    }, {});
  };

  const torihikisaki_table = groupBy(hiyo_table);
  for (let torihikisaki_table_row in torihikisaki_table) {
    const torihiki_saki = torihikisaki_table_row;
    let torihiki_nouki = [];
    torihikisaki_table[torihikisaki_table_row].forEach((v) => {
      let nounyu = new Date(v.nounyuyotei);
      let nounyunen = nounyu.getFullYear();
      let nounyutsuki = nounyu.getMonth() + 1;
      let nounyubi = nounyu.getDate();
      let nengappi = nounyunen + '年' + nounyutsuki + '月' + nounyubi + '日';
      torihiki_nouki.push(nengappi);
    });

    function tableRow_torihikisaki_cell(c_text) {
      const table_cell = new docx.TableCell({
        columnSpan: 1,
        children: [
          new docx.Paragraph({
            text: c_text,
            alignment: docx.AlignmentType.LEFT,
          }),
        ],
      });
      return table_cell;
    }

    const tableRow_torihikisaki = new docx.TableRow({
      children: [
        tableRow_torihikisaki_cell(torihiki_saki),
        tableRow_torihikisaki_cell(torihiki_nouki.join('、')),
      ],
    });
    table_torihiki.root.push(tableRow_torihikisaki);
  }

  //
  // 費用テーブルの作成
  //

  //
  // 費用テーブルのヘッダ作成
  //
  const table_hiyo = new docx.Table({
    rows: [
      new docx.TableRow({
        cantSplit: true,
        children: [
          new docx.TableCell({
            width: {
              size: 20,
              type: docx.WidthType.PERCENTAGE,
            },
            shading: {
              fill: 'F5F5F5',
            },
            verticalAlign: docx.VerticalAlign.CENTER,
            columnSpan: 1,
            rowSpan: 3,
            children: [
              new docx.Paragraph({
                text: '内容',
                style: 'Figure1',
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
          new docx.TableCell({
            width: {
              size: 7,
              type: docx.WidthType.PERCENTAGE,
            },
            shading: {
              fill: 'F5F5F5',
            },
            verticalAlign: docx.VerticalAlign.CENTER,
            columnSpan: 1,
            rowSpan: 3,
            children: [
              new docx.Paragraph({
                text: '決裁',
                style: 'Figure1',
                alignment: docx.AlignmentType.CENTER,
              }),
              new docx.Paragraph({
                text: '総額',
                style: 'Figure1',
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
          new docx.TableCell({
            width: {
              size: 14,
              type: docx.WidthType.PERCENTAGE,
            },
            shading: {
              fill: 'F5F5F5',
            },
            columnSpan: 2,
            rowSpan: 1,
            children: [
              new docx.Paragraph({
                text: 'FY2021',
                style: 'Figure1',
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
          new docx.TableCell({
            width: {
              size: 7,
              type: docx.WidthType.PERCENTAGE,
            },
            shading: {
              fill: 'F5F5F5',
            },
            columnSpan: 1,
            rowSpan: 1,
            children: [
              new docx.Paragraph({
                text: 'FY2021',
                style: 'Figure1',
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
          new docx.TableCell({
            width: {
              size: 7,
              type: docx.WidthType.PERCENTAGE,
            },
            shading: {
              fill: 'F5F5F5',
            },
            columnSpan: 1,
            rowSpan: 1,
            children: [
              new docx.Paragraph({
                text: 'FY2022',
                style: 'Figure1',
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
          new docx.TableCell({
            width: {
              size: 7,
              type: docx.WidthType.PERCENTAGE,
            },
            shading: {
              fill: 'F5F5F5',
            },
            columnSpan: 1,
            rowSpan: 1,
            children: [
              new docx.Paragraph({
                text: 'FY2023',
                style: 'Figure1',
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
          new docx.TableCell({
            width: {
              size: 7,
              type: docx.WidthType.PERCENTAGE,
            },
            shading: {
              fill: 'F5F5F5',
            },
            columnSpan: 1,
            rowSpan: 1,
            children: [
              new docx.Paragraph({
                text: 'FY2024',
                style: 'Figure1',
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
          new docx.TableCell({
            width: {
              size: 12,
              type: docx.WidthType.PERCENTAGE,
            },
            shading: {
              fill: 'F5F5F5',
            },
            verticalAlign: docx.VerticalAlign.CENTER,
            columnSpan: 1,
            rowSpan: 3,
            children: [
              new docx.Paragraph({
                text: '取引先',
                style: 'Figure1',
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
          new docx.TableCell({
            width: {
              size: 12,
              type: docx.WidthType.PERCENTAGE,
            },
            shading: {
              fill: 'F5F5F5',
            },
            verticalAlign: docx.VerticalAlign.CENTER,
            columnSpan: 1,
            rowSpan: 3,
            children: [
              new docx.Paragraph({
                text: '備考',
                style: 'Figure1',
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
        ],
      }),
      new docx.TableRow({
        cantSplit: true,
        children: [
          new docx.TableCell({
            shading: {
              fill: 'F5F5F5',
            },
            columnSpan: 2,
            rowSpan: 1,
            children: [
              new docx.Paragraph({
                text: 'IT開発費',
                style: 'Figure1',
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
          new docx.TableCell({
            shading: {
              fill: 'F5F5F5',
            },
            columnSpan: 4,
            rowSpan: 1,
            children: [
              new docx.Paragraph({
                text: 'IT運営費',
                style: 'Figure1',
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
        ],
      }),
      new docx.TableRow({
        cantSplit: true,
        children: [
          new docx.TableCell({
            width: {
              size: 7,
              type: docx.WidthType.PERCENTAGE,
            },
            shading: {
              fill: 'F5F5F5',
            },
            columnSpan: 1,
            rowSpan: 1,
            children: [
              new docx.Paragraph({
                text: '開発',
                style: 'Figure1',
                alignment: docx.AlignmentType.CENTER,
              }),
              new docx.Paragraph({
                text: '一時',
                style: 'Figure1',
                alignment: docx.AlignmentType.CENTER,
              }),
              new docx.Paragraph({
                text: '費用',
                style: 'Figure1',
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
          new docx.TableCell({
            width: {
              size: 7,
              type: docx.WidthType.PERCENTAGE,
            },
            shading: {
              fill: 'F5F5F5',
            },
            columnSpan: 1,
            rowSpan: 1,
            children: [
              new docx.Paragraph({
                text: '初年度',
                style: 'Figure1',
                alignment: docx.AlignmentType.CENTER,
              }),
              new docx.Paragraph({
                text: '運営',
                style: 'Figure1',
                alignment: docx.AlignmentType.CENTER,
              }),
              new docx.Paragraph({
                text: '費用',
                style: 'Figure1',
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
          new docx.TableCell({
            width: {
              size: 5,
              type: docx.WidthType.PERCENTAGE,
            },
            shading: {
              fill: 'F5F5F5',
            },
            columnSpan: 4,
            rowSpan: 1,
            children: [
              new docx.Paragraph({
                text: '年間費用',
                style: 'Figure1',
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
        ],
      }),
    ],
  });

  //
  // 費用テーブルの明細作成
  //
  function tableRow_hiyo_cell(p_text, p_alignment) {
    const table_cell = new docx.TableCell({
      columnSpan: 1,
      children: [
        new docx.Paragraph({
          text: p_text,
          style: 'Figure1',
          alignment: p_alignment,
        }),
      ],
    });
    return table_cell;
  }

  // 明細行作成
  hiyo_table.forEach((hiyo_table_row) => {
    const hiyo01 = String(
      Math.round(hiyo_table_row.value.hiyo01.value / 1000)
    ).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
    const hiyo02 = String(
      Math.round(hiyo_table_row.value.hiyo02.value / 1000)
    ).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
    const hiyo03_0 = String(
      Math.round(hiyo_table_row.value.hiyo03_0.value / 1000)
    ).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
    const hiyo03 = String(
      Math.round(hiyo_table_row.value.hiyo03.value / 1000)
    ).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
    const hiyo04 = String(
      Math.round(hiyo_table_row.value.hiyo04.value / 1000)
    ).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
    const hiyo05 = String(
      Math.round(hiyo_table_row.value.hiyo05.value / 1000)
    ).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
    const hiyo_kei = String(
      Math.round(hiyo_table_row.value.hiyo_kei.value / 1000)
    ).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
    const hiyo_biko = hiyo_table_row.value.hiyo_biko.value;
    const hiyo_naiyo = hiyo_table_row.value.hiyo_naiyo.value;
    const hiyo_torihikisaki = hiyo_table_row.value.hiyo_torihikisaki.value;

    const tableRow_hiyo = new docx.TableRow({
      cantSplit: true,
      children: [
        tableRow_hiyo_cell(hiyo_naiyo, docx.AlignmentType.LEFT),
        tableRow_hiyo_cell(hiyo_kei, docx.AlignmentType.RIGHT),
        tableRow_hiyo_cell(hiyo01, docx.AlignmentType.RIGHT),
        tableRow_hiyo_cell(hiyo02, docx.AlignmentType.RIGHT),
        tableRow_hiyo_cell(hiyo03_0, docx.AlignmentType.RIGHT),
        tableRow_hiyo_cell(hiyo03, docx.AlignmentType.RIGHT),
        tableRow_hiyo_cell(hiyo04, docx.AlignmentType.RIGHT),
        tableRow_hiyo_cell(hiyo05, docx.AlignmentType.RIGHT),
        tableRow_hiyo_cell(hiyo_torihikisaki, docx.AlignmentType.LEFT),
        tableRow_hiyo_cell(hiyo_biko, docx.AlignmentType.LEFT),
      ],
    });
    table_hiyo.root.push(tableRow_hiyo);
  });

  // 合計行作成
  const tableRow_hiyo_kei = new docx.TableRow({
    cantSplit: true,
    children: [
      tableRow_hiyo_cell('合計', docx.AlignmentType.LEFT),
      tableRow_hiyo_cell(hiyo_total, docx.AlignmentType.RIGHT),
      tableRow_hiyo_cell(it_kaihatsu_ichiji, docx.AlignmentType.RIGHT),
      tableRow_hiyo_cell(it_kaihatsu_uneihi, docx.AlignmentType.RIGHT),
      tableRow_hiyo_cell(it_uneihi_ty, docx.AlignmentType.RIGHT),
      tableRow_hiyo_cell(it_uneihi_ny, docx.AlignmentType.RIGHT),
      tableRow_hiyo_cell(it_uneihi_n2y, docx.AlignmentType.RIGHT),
      tableRow_hiyo_cell(it_uneihi_n3y, docx.AlignmentType.RIGHT),
      tableRow_hiyo_cell('', docx.AlignmentType.LEFT),
      tableRow_hiyo_cell('', docx.AlignmentType.LEFT),
    ],
  });

  table_hiyo.root.push(tableRow_hiyo_kei);

  //
  //添付ファイルテーブルの作成
  //

  let paragraph_files = new docx.Paragraph({
    alignment: docx.AlignmentType.LEFT,
    children: [],
  });
  if (file_mitsumori.length) {
    paragraph_files.root.push(
      new docx.TextRun({
        text:
          '・見積書、発注書　　　　　　　　　　' + file_mitsumori.length + '通',
      })
    );
  }

  if (file_torihikisaki_sentei.length) {
    paragraph_files.root.push(
      new docx.TextRun({
        break: true,
        text:
          '・取引先選定書　　　　　　　　　　　' +
          file_torihikisaki_sentei.length +
          '通',
      })
    );
  }

  if (file_sonota.length) {
    paragraph_files.root.push(
      new docx.TextRun({
        break: true,
        text:
          '・その他　　　　　　　　　　　　　　' + file_sonota.length + '通',
      })
    );
  }
  const figure_shiharai = [
    new docx.Paragraph({ children: paragraph_kessai_summary }),
    new docx.Paragraph({ text: '' }),
    new docx.Paragraph({
      text: '記',
      alignment: docx.AlignmentType.CENTER,
    }),
    new docx.Paragraph({ text: '' }),
    new docx.Paragraph({ text: '１．案件' }),
    new docx.Paragraph({ text: '　PJ No: ' + project_no }),
    new docx.Paragraph({ text: '　名称 : ' + project_name }),
    new docx.Paragraph({ text: '' }),
    new docx.Paragraph({
      text: '２．背景・経緯',
      children: paragraph_haikei,
    }),
    new docx.Paragraph({ text: '' }),
    new docx.Paragraph({
      text: '３．実施内容',
      children: paragraph_jisshi_naiyo,
    }),
    new docx.Paragraph({ text: '' }),
    new docx.Paragraph({ text: '４．発注先' }),
    table_torihiki,
    new docx.Paragraph({ text: '' }),
    new docx.Paragraph({ text: '５．費用' }),
    new docx.Paragraph({
      text: '（単位：千円）',
      alignment: docx.AlignmentType.RIGHT,
    }),
    table_hiyo,
    new docx.Paragraph({ text: '' }),
    new docx.Paragraph({ text: '６．添付資料' }),
    paragraph_files,
    new docx.Paragraph({ text: '' }),
    new docx.Paragraph({ text: '以上' }),
  ];
  // console.log(paragraph_kessai_summary);
  return figure_shiharai;
}
