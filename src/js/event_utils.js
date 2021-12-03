import * as word_format from './word_format';
import * as docx from 'docx';
import * as FileSaver from 'file-saver';

export function event_kessaitype(event) {
  let record = event.record;
  console.log('レコード情報', event);
  const kessaiSha = 'gcho';
  const footertext = 'テ　ス　ト　株　式　会　社　　　　　　　　　　　　　';

  var docxButton = document.createElement('button');
  docxButton.id = 'docx_button';
  docxButton.innerText = '決裁書Wordダウンロード';
  console.log(docxButton);
  document.body.appendChild(docxButton);

  // kintone.app.record.getHeaderMenuSpaceElement().appendChild(docxButton);
  docxButton.addEventListener('click', function () {
    docx.Packer.toBlob(
      word_format.createDoc(record, kessaiSha, footertext)
    ).then((blob) => {
      FileSaver.saveAs(blob, record.anken_name.value + '.docx');
      console.log('Document created successfully');
    });
  });

  return event;
}
