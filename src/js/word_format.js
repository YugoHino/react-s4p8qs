import * as docx from 'docx';

export function createDoc(record, kessaiSha,footertext) {
  let headertext = '';
  switch (kessaiSha) {
    case 'gcho':
      headertext = '決　裁　書（グループ長）';
      break;
    case 'bucho':
      headertext = '決　裁　書（部長）';
      break;
    case 'honbucho':
      headertext = '決　裁　書（本部長）';
      break;
  }
  const teian_busho = record.teian_busho.value.substr(2);
  const hakko_bango =
    record.FY.value + '-' + ('000' + record.record_no.value).slice(-3) + '号';
  const anken_name = record.anken_name.value;
  const anken_biko = record.anken_biko.value;
  const hiyo_total = String(Math.round(record.hiyo_total.value / 1000)).replace(
    /(\d)(?=(\d{3})+(?!\d))/g,
    '$1,'
  );
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
  const kessai_summary = record.kessai_summary.value;
  const project_no = record.project_no.value;
  const project_name = record.project_name.value;
  const system_no = record.system_no.value;
  const haikei = record.haikei.value;
  const jisshi_naiyo = record.jisshi_naiyo.value;
  const yosan_sochi = record.yosan_sochi.value;
  const yosan_sochi_count = yosan_sochi.length;
  const yosan_keijo = record.yosan_keijo.value.substr(2);
  const yosan_riyu = record.yosan_riyu.value;
  const hiyo_table = record.hiyo_table.value;
  const file_mitsumori = record.file_mitsumori.value;
  const file_torihikisaki_sentei = record.file_torihikisaki_sentei.value;
  const file_sonota = record.file_sonota.value;

  const figure_common = new docx.Table({
    margins: {
      top: 10,
      bottom: 10,
      left: 10,
      right: 10,
    },
    rows: [
      new docx.TableRow({
        children: [
          new docx.TableCell({
            width: {
              size: 5,
              type: docx.WidthType.PERCENTAGE,
            },
            columnSpan: 1,
            children: [
              new docx.Paragraph({
                text: '提',
                alignment: docx.AlignmentType.CENTER,
              }),
              new docx.Paragraph({
                text: '案',
                alignment: docx.AlignmentType.CENTER,
              }),
              new docx.Paragraph({
                text: '部',
                alignment: docx.AlignmentType.CENTER,
              }),
              new docx.Paragraph({
                text: '署',
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
          new docx.TableCell({
            width: {
              size: 65,
              type: docx.WidthType.PERCENTAGE,
            },
            columnSpan: 5,
            verticalAlign: docx.VerticalAlign.CENTER,
            children: [
              new docx.Paragraph({
                text: teian_busho,
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
          new docx.TableCell({
            width: {
              size: 5,
              type: docx.WidthType.PERCENTAGE,
            },
            columnSpan: 1,
            children: [
              new docx.Paragraph({
                text: '発',
                alignment: docx.AlignmentType.CENTER,
              }),
              new docx.Paragraph({
                text: '行',
                alignment: docx.AlignmentType.CENTER,
              }),
              new docx.Paragraph({
                text: '番',
                alignment: docx.AlignmentType.CENTER,
              }),
              new docx.Paragraph({
                text: '号',
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
          new docx.TableCell({
            width: {
              size: 25,
              type: docx.WidthType.PERCENTAGE,
            },
            columnSpan: 2,
            verticalAlign: docx.VerticalAlign.CENTER,
            children: [
              new docx.Paragraph({
                text: hakko_bango,
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
        ],
      }),
      new docx.TableRow({
        children: [
          new docx.TableCell({
            width: {
              size: 5,
              type: docx.WidthType.PERCENTAGE,
            },
            columnSpan: 1,
            children: [
              new docx.Paragraph({
                text: '',
                alignment: docx.AlignmentType.CENTER,
              }),
              new docx.Paragraph({
                text: '件',
                alignment: docx.AlignmentType.CENTER,
              }),
              new docx.Paragraph({
                text: '名',
                alignment: docx.AlignmentType.CENTER,
              }),
              new docx.Paragraph({
                text: '',
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
          new docx.TableCell({
            width: {
              size: 95,
              type: docx.WidthType.PERCENTAGE,
            },
            columnSpan: 8,
            verticalAlign: docx.VerticalAlign.CENTER,
            children: [
              new docx.Paragraph({
                text: anken_name,
                alignment: docx.AlignmentType.LEFT,
              }),
            ],
          }),
        ],
      }),
      new docx.TableRow({
        children: [
          new docx.TableCell({
            width: {
              size: 25,
              type: docx.WidthType.PERCENTAGE,
            },
            columnSpan: 3,
            children: [
              new docx.Paragraph({
                text: '決裁総額（税抜）',
                alignment: docx.AlignmentType.LEFT,
              }),
            ],
          }),
          new docx.TableCell({
            width: {
              size: 50,
              type: docx.WidthType.PERCENTAGE,
            },
            columnSpan: 5,
            children: [
              new docx.Paragraph({
                text: hiyo_total + '千円',
                alignment: docx.AlignmentType.LEFT,
              }),
            ],
          }),
          new docx.TableCell({
            width: {
              size: 25,
              type: docx.WidthType.PERCENTAGE,
            },
            columnSpan: 1,
            rowSpan: 3 + yosan_sochi_count,
            children: [
              new docx.Paragraph({
                text: '（備考）',
                alignment: docx.AlignmentType.CENTER,
              }),
              new docx.Paragraph({
                text: anken_biko,
                alignment: docx.AlignmentType.LEFT,
              }),
            ],
          }),
        ],
      }),
      new docx.TableRow({
        children: [
          new docx.TableCell({
            width: {
              size: 80,
              type: docx.WidthType.PERCENTAGE,
            },
            columnSpan: 8,
            children: [
              new docx.Paragraph({
                text: '本年度　予算措置',
                alignment: docx.AlignmentType.LEFT,
              }),
            ],
          }),
        ],
      }),
      new docx.TableRow({
        children: [
          new docx.TableCell({
            width: {
              size: 20,
              type: docx.WidthType.PERCENTAGE,
            },
            columnSpan: 2,
            children: [
              new docx.Paragraph({
                text: '予算区分',
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
          new docx.TableCell({
            width: {
              size: 20,
              type: docx.WidthType.PERCENTAGE,
            },
            columnSpan: 2,
            children: [
              new docx.Paragraph({
                text: '予算科目',
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
          new docx.TableCell({
            width: {
              size: 20,
              type: docx.WidthType.PERCENTAGE,
            },
            columnSpan: 1,
            children: [
              new docx.Paragraph({
                text: '金額（税抜き）',
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
          new docx.TableCell({
            width: {
              size: 15,
              type: docx.WidthType.PERCENTAGE,
            },
            verticalAlign: docx.VerticalAlign.CENTER,
            columnSpan: 3,
            rowSpan: 1 + yosan_sochi_count,
            children: [
              new docx.Paragraph({
                text: yosan_keijo,
                alignment: docx.AlignmentType.CENTER,
              }),
            ],
          }),
        ],
      }),
    ],
  });

  yosan_sochi.forEach((yosan_sochi_row) => {
    const yosan_kubun = yosan_sochi_row.value.yosan_kubun.value;
    const yosan_kamoku = yosan_sochi_row.value.yosan_kamoku.value;
    const yosan_kingaku =
      String(
        Math.round(yosan_sochi_row.value.yosan_kingaku.value / 1000)
      ).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,') + '千円';
    const tableRow_yosan_sochi = new docx.TableRow({
      children: [
        new docx.TableCell({
          width: {
            size: 10,
            type: docx.WidthType.PERCENTAGE,
          },
          columnSpan: 2,
          children: [
            new docx.Paragraph({
              text: yosan_kubun,
              alignment: docx.AlignmentType.LEFT,
            }),
          ],
        }),
        new docx.TableCell({
          width: {
            size: 20,
            type: docx.WidthType.PERCENTAGE,
          },
          columnSpan: 2,
          children: [
            new docx.Paragraph({
              text: yosan_kamoku,
              alignment: docx.AlignmentType.LEFT,
            }),
          ],
        }),
        new docx.TableCell({
          width: {
            size: 20,
            type: docx.WidthType.PERCENTAGE,
          },
          columnSpan: 1,
          children: [
            new docx.Paragraph({
              text: yosan_kingaku,
              alignment: docx.AlignmentType.RIGHT,
            }),
          ],
        }),
        new docx.TableCell({
          width: {
            size: 20,
            type: docx.WidthType.PERCENTAGE,
          },
          columnSpan: 3,
          verticalMerge: 'continue',
          children: [
            new docx.Paragraph({
              text: '',
              alignment: docx.AlignmentType.CENTER,
            }),
          ],
        }),
        new docx.TableCell({
          width: {
            size: 20,
            type: docx.WidthType.PERCENTAGE,
          },
          columnSpan: 1,
          verticalMerge: 'continue',
          children: [
            new docx.Paragraph({
              text: '',
              alignment: docx.AlignmentType.CENTER,
            }),
          ],
        }),
      ],
    });
    figure_common.root.push(tableRow_yosan_sochi);
  });

  const row_yosan_riyu = new docx.TableRow({
    children: [
      new docx.TableCell({
        width: {
          size: 100,
          type: docx.WidthType.PERCENTAGE,
        },
        columnSpan: 9,
        children: [
          new docx.Paragraph({
            text: '（予算金額超過、予算未計上の理由、対応および影響）',
            alignment: docx.AlignmentType.LEFT,
          }),
          new docx.Paragraph({
            text: yosan_riyu,
            alignment: docx.AlignmentType.LEFT,
          }),
          new docx.Paragraph({ text: '', alignment: docx.AlignmentType.LEFT }),
        ],
      }),
    ],
  });

  figure_common.root.push(row_yosan_riyu);

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
  function table_torihiki_cell(c_size,c_text){
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
    }),
    return table_cell
  }

  const table_torihiki = new docx.Table({
    rows: [
      new docx.TableRow({
        children: [
          table_torihiki_cell(40,'取引先'),
          table_torihiki_cell(60,'納入予定日'),
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

    function tableRow_torihikisaki_cell(c_text){
      const table_cell = new docx.TableCell({
        columnSpan: 1,
        children: [
          new docx.Paragraph({
            text: c_text,
            alignment: docx.AlignmentType.LEFT,
          }),
        ],
      }),
      return table_cell
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
  function tableRow_hiyo_cell(p_text,p_alignment){
    const table_cell = new docx.TableCell({
      columnSpan: 1,
      children: [
        new docx.Paragraph({
          text: p_text,
          style: 'Figure1',
          alignment: p_alignment,
        }),
      ],
    }),
    return table_cell
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
        tableRow_hiyo_cell(hiyo_naiyo,docx.AlignmentType.LEFT),
        tableRow_hiyo_cell(hiyo_kei,docx.AlignmentType.RIGHT),
        tableRow_hiyo_cell(hiyo01,docx.AlignmentType.RIGHT),
        tableRow_hiyo_cell(hiyo02,docx.AlignmentType.RIGHT),
        tableRow_hiyo_cell(hiyo03_0,docx.AlignmentType.RIGHT),
        tableRow_hiyo_cell(hiyo03,docx.AlignmentType.RIGHT),
        tableRow_hiyo_cell(hiyo04,docx.AlignmentType.RIGHT),
        tableRow_hiyo_cell(hiyo05,docx.AlignmentType.RIGHT),
        tableRow_hiyo_cell(hiyo_torihikisaki,docx.AlignmentType.LEFT),
        tableRow_hiyo_cell(hiyo_biko,docx.AlignmentType.LEFT),
      ],
    });
    table_hiyo.root.push(tableRow_hiyo);
  });

  // 合計行作成
  const tableRow_hiyo_kei = new docx.TableRow({
    cantSplit: true,
    children: [
      tableRow_hiyo_cell('合計',docx.AlignmentType.LEFT),
      tableRow_hiyo_cell(hiyo_total,docx.AlignmentType.RIGHT),
      tableRow_hiyo_cell(it_kaihatsu_ichiji,docx.AlignmentType.RIGHT),
      tableRow_hiyo_cell(it_kaihatsu_uneihi,docx.AlignmentType.RIGHT),
      tableRow_hiyo_cell(it_uneihi_ty,docx.AlignmentType.RIGHT),
      tableRow_hiyo_cell(it_uneihi_ny,docx.AlignmentType.RIGHT),
      tableRow_hiyo_cell(it_uneihi_n2y,docx.AlignmentType.RIGHT),
      tableRow_hiyo_cell(it_uneihi_n3y,docx.AlignmentType.RIGHT),
      tableRow_hiyo_cell('',docx.AlignmentType.LEFT),
      tableRow_hiyo_cell('',docx.AlignmentType.LEFT),
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



  // 
  // 出力するコンテンツをパターンから選択
  // 
  function document_contents(){

    const document_content =[
      figure_common,
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
      new docx.Paragraph({ text: '以上' })
    ]

    return document_content
  }

  // 
  //本体作成
  // 
  const doc = new docx.Document({
    styles: {
      paragraphStyles: [
        {
          name: 'Normal',
          run: {
            size: 24,
            font: 'ＭＳ 明朝',
          },
          paragraph: {
            spacing: {
              before: 10,
            },
            indent: {
              left: 50,
              right: 50,
            },
          },
        },
        {
          name: 'Figure1',
          run: {
            size: 18,
            font: 'ＭＳ 明朝',
          },
          paragraph: {
            spacing: {
              before: 10,
            },
            indent: {
              left: 50,
              right: 50,
            },
          },
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            borders: {
              pageBorderLeft: {
                style: docx.BorderStyle.SINGLE,
                size: 15,
                color: '000000',
                space: 0,
              },
              pageBorderRight: {
                style: docx.BorderStyle.SINGLE,
                size: 15,
                color: '000000',
                space: 0,
              },
              pageBorderTop: {
                style: docx.BorderStyle.SINGLE,
                size: 15,
                color: 'auto',
                space: 0,
              },
              pageBorderBottom: {
                style: docx.BorderStyle.SINGLE,
                size: 15,
                color: 'auto',
                space: 0,
              },
            },
          },
        },
        headers: {
          default: new docx.Header({
            // The standard default header
            children: [
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,

                children: [
                  new docx.TextRun({
                    text: headertext,
                    size: 48,
                    bold: true,
                  }),
                ],
              }),
            ],
          }),
          first: new docx.Header({
            // The first header
            children: [],
          }),
          even: new docx.Header({
            // The header on every other page
            children: [],
          }),
        },
        footers: {
          default: new docx.Footer({
            children: [
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                text: '',
                children: [
                  new docx.TextRun({ text: footertext }),
                  new docx.TextRun({ text: 'No.' }),
                  new docx.TextRun({
                    children: [
                      '         ',
                      docx.PageNumber.CURRENT,
                      '/',
                      docx.PageNumber.TOTAL_PAGES,
                    ],
                    underline: { color: '000000' },
                  }),
                ],
              }),
            ],
          }),
        },

        children: document_contents()
      },
    ],
  });
  return doc;
}
