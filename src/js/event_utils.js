import * as word_format from './word_format';
import * as docx from 'docx';
import * as FileSaver from 'file-saver';

export function event_kessaitype(event) {
  let record = event.record;
  console.log('レコード情報', event);
  const kessaiSha = 'gcho';
  // let return_setFieldShown = setFieldShown(event);
  // let kessaiSha = return_setFieldShown[0];
  // let hidden_fields = return_setFieldShown[1];

  const docxDLButtonId = 'kessai_download';
  // const div1 = document.getElementById('div1');
  // let form = document.createElement('form');
  // form.id = 'Form';
  var docxButton = document.createElement('button');
  // let docxButton = document.getElementById('button1');
  // let docxButton = new kintoneUIComponent.Button({ text: '決裁' });
  docxButton.id = 'docx_button';
  docxButton.innerText = '決裁書Wordダウンロード';
  console.log(docxButton);
  // form.appendChild(docxButton);
  // let element = document.getElementsByTagName('p');
  // element.appendChild(docxButton);
  document.body.appendChild(docxButton);

  // kintone.app.record.getHeaderMenuSpaceElement().appendChild(docxButton);
  docxButton.addEventListener('click', function () {
    docx.Packer.toBlob(word_format.createDoc(record, kessaiSha)).then(
      (blob) => {
        FileSaver.saveAs(blob, record.anken_name.value + '.docx');
        console.log('Document created successfully');
      }
    );
  });

  return event;
}
