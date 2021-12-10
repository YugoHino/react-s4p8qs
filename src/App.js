import React from 'react';
import './style.css';
import * as event_utils from './js/event_utils';

export default function App() {
  // let events = {};
  return (
    <div id="div1">
      <h1>Hello StackBlitz!</h1>
      <p>Start editing to see some magic happen :)</p>
    </div>
  );
}
(() => {
  // let events = {};
  const events = {
    record: {
      record_no: { value: 1 },
      teian_busho: { value: '1.企画部企画グル' },
      FY: { value: '21' },
      record_no: { value: 11 },
      anken_name: { value: 'test_anken' },
      anken_biko: {
        value:
          '備考ですあああああああああああああああああああああああああああああああああああああああああ',
      },
      // kessai_type: { value: '1.IT関係費' },
      kessai_type: { value: '2.業務推進費' },
      hiyo_total: { value: 1470000 },
      it_kaihatsu_ichiji: { value: 10000 },
      it_kaihatsu_uneihi: { value: 20000 },
      it_uneihi_ty: { value: 200000 },
      it_uneihi_ny: { value: 20000 },
      it_uneihi_n2y: { value: 30000 },
      it_uneihi_n3y: { value: 40000 },
      kessai_summary: { value: '決裁サマリてすとの文章です' },
      project_no: { value: 'pjnoxxx/xxx' },
      project_name: { value: '統合データ分析基盤' },
      system_no: { value: 'sysno/xxx' },
      haikei: { value: 'このプロジェクトの背景です' },
      jisshi_naiyo: {
        value:
          'GitHub Enterpriseライセンスの購入\n・9月 15ライセンス（Enterprise/Organization管理者分）\n・10月 50ライセンス（研究開発本部、製薬技術本部分）',
      },
      file_mitsumori: {
        value: [
          {
            name: 'test.docx',
          },
          {
            name: 'test1.docx',
          },
        ],
      },
      file_torihikisaki_sentei: {
        value: [
          {
            name: 'torihiki.docx',
          },
        ],
      },
      file_sonota: {
        value: [],
      },
      yosan_sochi: {
        value: [
          {
            value: {
              yosan_kamoku: {
                value: 'IT開発費',
              },
              yosan_kingaku: {
                value: '80',
              },
              yosan_kubun: {
                value: 'その他経費',
              },
            },
          },
          {
            value: {
              yosan_kamoku: {
                value: 'IT開発費2',
              },
              yosan_kingaku: {
                value: '20',
              },
              yosan_kubun: {
                value: 'その他経費2',
              },
            },
          },
        ],
      },
      yosan_keijo: {
        value: '3.予算未計上',
      },
      yosan_riyu: {
        value: '',
      },
      gyomusuishin_hiyo_table: {
        value: [
          {
            value: {
              gyomusuishin_hiyo_kei: {
                value: '9000000',
              },
              gyomusuishin_hiyo_naiyo: {
                value: 'とりひき',
              },
              gyomusuishin_hiyo_seikyu: {
                value: '2021-12-14',
              },
              gyomusuishin_hiyo_nounyuyotei: {
                value: '2021-12-15',
              },
              gyomusuishin_hiyo_torihikisaki: {
                value: 'test',
              },
              gyomusuishin_hiyo_kamoku: {
                value: 'IT直接費',
              },
              gyomusuishin_hiyo_biko: {
                value: 'コメントはこちらにはいってきます',
              },
            },
          },
          {
            value: {
              gyomusuishin_hiyo_kei: {
                value: '100000',
              },
              gyomusuishin_hiyo_naiyo: {
                value: 'とりひき2',
              },
              gyomusuishin_hiyo_seikyu: {
                value: '2021-12-16',
              },
              gyomusuishin_hiyo_nounyuyotei: {
                value: '2021-12-16',
              },
              gyomusuishin_hiyo_torihikisaki: {
                value: 'test2',
              },
              gyomusuishin_hiyo_kamoku: {
                value: 'IT直接費２',
              },
              gyomusuishin_hiyo_biko: {
                value: 'コメントはこちらにはいってきます２',
              },
            },
          },
        ],
      },
      hiyo_table: {
        value: [
          {
            value: {
              hiyo01: {
                value: '100000',
              },
              hiyo02: {
                value: '200',
              },
              hiyo03_0: {
                value: '310',
              },
              hiyo03: {
                value: '300',
              },
              hiyo04: {
                value: '400',
              },
              hiyo05: {
                value: '500',
              },
              hiyo_biko: {
                value: 'テスト備考',
              },
              hiyo_nounyuyotei: {
                value: '2021-01-01',
              },
              hiyo_kei: {
                value: '1000',
              },
              hiyo_naiyo: {
                value: 'ライセンス2',
              },
              hiyo_torihikisaki: {
                value: '企画株式会社A',
              },
            },
          },
          {
            value: {
              hiyo01: {
                value: '100',
              },
              hiyo02: {
                value: '200',
              },
              hiyo03_0: {
                value: '310',
              },
              hiyo03: {
                value: '300',
              },
              hiyo04: {
                value: '400',
              },
              hiyo05: {
                value: '500',
              },
              hiyo_biko: {
                value: 'テスト備考',
              },
              hiyo_nounyuyotei: {
                value: '2021-08-01',
              },
              hiyo_kei: {
                value: '1000',
              },
              hiyo_naiyo: {
                value: 'ライセンス2',
              },
              hiyo_torihikisaki: {
                value: '企画株式会社A',
              },
            },
          },
          {
            value: {
              hiyo01: {
                value: '100',
              },
              hiyo02: {
                value: '200',
              },
              hiyo03_0: {
                value: '310',
              },
              hiyo03: {
                value: '300',
              },
              hiyo04: {
                value: '400',
              },
              hiyo05: {
                value: '500',
              },
              hiyo_biko: {
                value: 'テスト備考',
              },
              hiyo_nounyuyotei: {
                value: '2021-07-01',
              },
              hiyo_kei: {
                value: '1000',
              },
              hiyo_naiyo: {
                value: 'ライセンス１',
              },
              hiyo_torihikisaki: {
                value: '企画株式会社A',
              },
            },
          },
          {
            value: {
              hiyo01: {
                value: '100',
              },
              hiyo02: {
                value: '200',
              },
              hiyo03_0: {
                value: '310',
              },
              hiyo03: {
                value: '300',
              },
              hiyo04: {
                value: '400',
              },
              hiyo05: {
                value: '500',
              },
              hiyo_biko: {
                value: 'テスト備考',
              },
              hiyo_nounyuyotei: {
                value: '2021-07-01',
              },
              hiyo_kei: {
                value: '1000',
              },
              hiyo_naiyo: {
                value: 'ライセンス3',
              },
              hiyo_torihikisaki: {
                value: '企画株式会社b',
              },
            },
          },
        ],
      },
    },
  };

  console.log('ok');
  event_utils.event_kessaitype(events);
})();
