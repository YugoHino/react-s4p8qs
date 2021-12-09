import * as docx from 'docx';
import * as word_format_common from './word_format_common';
import * as word_format_shiharai from './word_format_shiharai';

export function createDoc(record, kessaiSha, footertext) {
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

  //
  // 出力するコンテンツをパターンから選択
  //
  function document_contents() {
    const document_content = [
      word_format_common.getFigureCommon(record),
      // new docx.Paragraph({ children: paragraph_kessai_summary }),
      // new docx.Paragraph({ text: '' }),
      // new docx.Paragraph({
      //   text: '記',
      //   alignment: docx.AlignmentType.CENTER,
      // }),
      // new docx.Paragraph({ text: '' }),
      // new docx.Paragraph({ text: '１．案件' }),
      // new docx.Paragraph({ text: '　PJ No: ' + project_no }),
      // new docx.Paragraph({ text: '　名称 : ' + project_name }),
      // new docx.Paragraph({ text: '' }),
      // new docx.Paragraph({
      //   text: '２．背景・経緯',
      //   children: paragraph_haikei,
      // }),
      // new docx.Paragraph({ text: '' }),
      // new docx.Paragraph({
      //   text: '３．実施内容',
      //   children: paragraph_jisshi_naiyo,
      // }),
      // new docx.Paragraph({ text: '' }),
      // new docx.Paragraph({ text: '４．発注先' }),
      // table_torihiki,
      // new docx.Paragraph({ text: '' }),
      // new docx.Paragraph({ text: '５．費用' }),
      // new docx.Paragraph({
      //   text: '（単位：千円）',
      //   alignment: docx.AlignmentType.RIGHT,
      // }),
      // table_hiyo,
      // new docx.Paragraph({ text: '' }),
      // new docx.Paragraph({ text: '６．添付資料' }),
      // paragraph_files,
      // new docx.Paragraph({ text: '' }),
      // new docx.Paragraph({ text: '以上' }),
    ];

    return document_content;
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

        children: document_contents(),
      },
    ],
  });
  return doc;
}
