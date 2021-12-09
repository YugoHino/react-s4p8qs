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
    let document_content = [word_format_common.getFigureCommon(record)];
    console.log(word_format_shiharai.getFigureShiharai(record));
    document_content.concat(word_format_shiharai.getFigureShiharai(record));
    // console.log(document_content[1]);
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
