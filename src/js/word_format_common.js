import * as docx from 'docx';

export function getFigureCommon(record) {
  const teian_busho = record.teian_busho.value.substr(2);
  const hakko_bango =
    record.FY.value + '-' + ('000' + record.record_no.value).slice(-3) + '号';
  const anken_name = record.anken_name.value;
  const anken_biko = record.anken_biko.value;
  const hiyo_total = String(Math.round(record.hiyo_total.value / 1000)).replace(
    /(\d)(?=(\d{3})+(?!\d))/g,
    '$1,'
  );
  const yosan_sochi = record.yosan_sochi.value;
  const yosan_sochi_count = yosan_sochi.length;
  const yosan_keijo = record.yosan_keijo.value.substr(2);
  const yosan_riyu = record.yosan_riyu.value;

  //
  // 共通表の作成
  //

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

  return figure_common;
}
