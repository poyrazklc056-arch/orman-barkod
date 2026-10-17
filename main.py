# -*- coding: utf-8 -*-
"""Orman Barkod Kayıt - Android/Kivy sürümü"""
import json, math, os
from datetime import datetime
from pathlib import Path

from kivy.app import App
from kivy.lang import Builder
from kivy.metrics import dp
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.spinner import Spinner
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.popup import Popup

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    Workbook = None

BOYLAR = ['1','1,20','1,25','1,30','1,40','1,50','2','2,50','3','3,50','4','5','6','7','8','9','10','11','12']
CINSLER = ['Sarıçam','Göknar','Karaçam','Kayın','Meşe','Ladin','Diğer']
URUNLER = ['Tomruk','Kağıtlık odun','Sanayi odunu','Maden direği','Tel direği']

KV = '''
<Label>:
    color: 0.09, 0.20, 0.23, 1
    size_hint_y: None
    height: dp(28)
    text_size: self.width, None
    halign: 'left'
    valign: 'middle'
<TextInput>:
    size_hint_y: None
    height: dp(44)
    multiline: False
    padding: dp(10), dp(10)
<Spinner>:
    size_hint_y: None
    height: dp(44)
'''

class OrmanBarkodAndroid(App):
    def build(self):
        self.title = 'Orman Barkod Kayıt'
        self.kayitlar = []
        self.boylar = BOYLAR.copy()
        self.duzenlenen_index = None
        self.veri_dosyasi = os.path.join(self.user_data_dir, 'OrmanBarkodKayitlari.json')
        self.root_box = BoxLayout(orientation='vertical', padding=dp(12), spacing=dp(8))
        self.root_box.add_widget(Label(text='🌲 ORMAN BARKOD KAYIT', size_hint_y=None, height=dp(46), font_size='21sp', bold=True, color=(0.08,0.28,0.28,1)))
        nav = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(8))
        nav.add_widget(self.btn('KAYIT GİRİŞİ', self.giris_goster))
        nav.add_widget(self.btn('KAYITLAR', self.kayitlar_goster))
        self.root_box.add_widget(nav)
        self.body = BoxLayout(orientation='vertical', spacing=dp(8))
        self.root_box.add_widget(self.body)
        self.ozet = Label(text='Kayıt: 0 | Adet: 0 | Hacim: 0,000 m³', size_hint_y=None, height=dp(34), bold=True, color=(0.08,0.38,0.31,1))
        self.root_box.add_widget(self.ozet)
        self._yukle()
        self.giris_goster()
        self.ozet_guncelle()
        return self.root_box

    def btn(self, text, fn):
        b = Button(text=text, size_hint_y=None, height=dp(44), background_normal='', background_color=(0.09,0.45,0.38,1), color=(1,1,1,1))
        b.bind(on_release=lambda *_: fn())
        return b

    def alan(self, parent, label, default=''):
        parent.add_widget(Label(text=label, size_hint_y=None, height=dp(24), bold=True))
        t = TextInput(text=str(default), hint_text=label)
        parent.add_widget(t)
        return t

    def secim(self, parent, label, values, default):
        parent.add_widget(Label(text=label, size_hint_y=None, height=dp(24), bold=True))
        s = Spinner(text=default, values=values)
        parent.add_widget(s)
        return s

    def giris_goster(self):
        self.body.clear_widgets()
        sc = ScrollView()
        form = BoxLayout(orientation='vertical', spacing=dp(6), size_hint_y=None, padding=(0,0,0,dp(12)))
        form.bind(minimum_height=form.setter('height'))
        self.barkod = self.alan(form, 'Barkod (okut veya yaz)')
        self.cins = self.secim(form, 'Ağaç cinsi', CINSLER, CINSLER[0])
        self.urun = self.secim(form, 'Ürün', URUNLER, URUNLER[0])
        self.boy = self.secim(form, 'Boy (metre)', self.uygun_boylar(URUNLER[0]), '3')
        self.cap = self.alan(form, 'Çap (cm)')
        self.adet = self.alan(form, 'Adet', '1')
        self.aciklama = self.alan(form, 'Açıklama / kişi')
        self.hacim_label = Label(text='Hesaplanan hacim: —', size_hint_y=None, height=dp(34), bold=True)
        form.add_widget(self.hacim_label)
        self.urun.bind(text=self.urun_degisti)
        form.add_widget(self.btn('HACMİ HESAPLA', self.hacim_hesapla))
        self.kaydet_btn = self.btn('KAYDI EKLE', self.ekle)
        form.add_widget(self.kaydet_btn)
        form.add_widget(self.btn('BOY LİSTESİNİ DÜZENLE', self.boy_listesi_popup))
        sc.add_widget(form); self.body.add_widget(sc)

    def uygun_boylar(self, urun):
        out=[]
        for b in self.boylar:
            try: n=float(b.replace(',','.'))
            except Exception: continue
            if n < 1.5 and urun != 'Sanayi odunu': continue
            if n > 3.5 and not n.is_integer(): continue
            out.append(b)
        return out or ['1,50','2','3','3,50','4','5','6']

    def urun_degisti(self, spinner, value):
        if hasattr(self, 'boy'):
            vals=self.uygun_boylar(value); self.boy.values=vals
            if self.boy.text not in vals: self.boy.text=next((x for x in vals if float(x.replace(',','.'))>=1.5), vals[0])

    def hacim_hesapla(self):
        try:
            cap=float(self.cap.text.strip().replace(',','.')); boy=float(self.boy.text.replace(',','.')); adet=int(self.adet.text or '1')
            if cap<=0 or boy<=0 or adet<=0: raise ValueError()
            v=math.pi*(cap/200)**2*boy*adet
            self.hacim_label.text=f'Hesaplanan hacim: {v:.3f} m³'.replace('.',',')
        except Exception:
            self.hacim_label.text='Hacim için çap, boy ve adet bilgilerini kontrol edin.'

    def popup(self, title, message):
        box=BoxLayout(orientation='vertical', padding=dp(12), spacing=dp(10))
        box.add_widget(Label(text=message))
        p=Popup(title=title, content=box, size_hint=(.88,.42))
        box.add_widget(Button(text='Tamam', size_hint_y=None, height=dp(44), on_release=p.dismiss)); p.open()

    def ekle(self):
        barkod=self.barkod.text.strip()
        if not barkod: self.popup('Barkod gerekli','Barkod okutun veya yazın.'); return
        try:
            cap=float(self.cap.text.strip().replace(',','.')); boy=float(self.boy.text.replace(',','.')); adet=int(self.adet.text or '1')
            if cap<=0 or boy<=0 or adet<1: raise ValueError()
        except Exception:
            self.popup('Bilgi hatası','Çap ve boy sıfırdan büyük, adet en az 1 olmalı.'); return
        for i,r in enumerate(self.kayitlar):
            if str(r['Barkod']).casefold()==barkod.casefold() and i != self.duzenlenen_index:
                self.popup('Tekrarlanan barkod','Bu barkod daha önce kaydedilmiş.'); return
        birim=math.pi*(cap/200)**2*boy
        rec={'Barkod':barkod,'Ağaç cinsi':self.cins.text,'Ürün':self.urun.text,'Boy (m)':boy,'Çap (cm)':cap,'Adet':adet,'Birim hacim (m³)':birim,'Toplam hacim (m³)':birim*adet,'Açıklama':self.aciklama.text.strip(),'Kayıt zamanı':datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
        if self.duzenlenen_index is None: self.kayitlar.append(rec)
        else: self.kayitlar[self.duzenlenen_index]=rec; self.duzenlenen_index=None
        self._yedekle(); self.ozet_guncelle(); self.popup('Kaydedildi','Kayıt uygulamanın içine kaydedildi.')
        self.barkod.text=''; self.cap.text=''; self.adet.text='1'; self.aciklama.text=''; self.kaydet_btn.text='KAYDI EKLE'

    def ozet_guncelle(self):
        self.ozet.text=f"Kayıt: {len(self.kayitlar)} | Adet: {sum(r['Adet'] for r in self.kayitlar)} | Hacim: {sum(r['Toplam hacim (m³)'] for r in self.kayitlar):.3f} m³".replace('.',',')

    def kayitlar_goster(self):
        self.body.clear_widgets()
        bar=BoxLayout(size_hint_y=None,height=dp(44),spacing=dp(6))
        bar.add_widget(self.btn('EXCEL’E AKTAR', self.excele)); bar.add_widget(self.btn('SEÇİLİ KAYDI SİL', self.secili_sil))
        self.body.add_widget(bar)
        sc=ScrollView(); grid=GridLayout(cols=1, spacing=dp(5), size_hint_y=None, padding=(0,0,0,dp(10))); grid.bind(minimum_height=grid.setter('height'))
        if not self.kayitlar: grid.add_widget(Label(text='Henüz kayıt yok.', size_hint_y=None, height=dp(50)))
        for i,r in enumerate(self.kayitlar):
            line=(f"{i+1}. Barkod: {r['Barkod']}\n{r['Ağaç cinsi']} | {r['Ürün']} | Boy {r['Boy (m)']} m | Çap {r['Çap (cm)']} cm\nAdet: {r['Adet']} | Hacim: {r['Toplam hacim (m³)']:.3f} m³ | {r['Açıklama']}")
            row=BoxLayout(size_hint_y=None,height=dp(102),spacing=dp(5))
            row.add_widget(Label(text=line, halign='left', valign='middle', text_size=(None,None)))
            row.add_widget(self.btn('DÜZENLE', lambda ix=i:self.kaydi_duzenle(ix)))
            grid.add_widget(row)
        sc.add_widget(grid); self.body.add_widget(sc)

    def kaydi_duzenle(self, idx):
        r=self.kayitlar[idx]; self.duzenlenen_index=idx; self.giris_goster()
        self.barkod.text=r['Barkod']; self.cins.text=r['Ağaç cinsi']; self.urun.text=r['Ürün']; self.urun_degisti(self.urun,self.urun.text)
        self.boy.text=str(r['Boy (m)']).replace('.',','); self.cap.text=str(r['Çap (cm)']).replace('.',','); self.adet.text=str(r['Adet']); self.aciklama.text=r['Açıklama']; self.kaydet_btn.text='DEĞİŞİKLİĞİ KAYDET'; self.hacim_hesapla()

    def secili_sil(self):
        self.popup('Silme','Kayıtları tek tek DÜZENLE ekranından değil, listedeki kayıt için uygulama menüsünden silme desteği bu sürümde sınırlıdır. JSON yedeğiniz korunur.')

    def boy_listesi_popup(self):
        box=BoxLayout(orientation='vertical',padding=dp(10),spacing=dp(8)); ti=TextInput(text='\n'.join(self.boylar),multiline=True)
        box.add_widget(Label(text='Her boyu ayrı satıra yazın (metre).',size_hint_y=None,height=dp(28))); box.add_widget(ti)
        p=Popup(title='Boy listesini düzenle',content=box,size_hint=(.9,.8))
        def save(*_):
            vals=[x.strip() for x in ti.text.splitlines() if x.strip()]
            try:
                for x in vals:
                    n=float(x.replace(',','.'))
                    if n<=0: raise ValueError()
                if not vals: raise ValueError()
            except Exception: self.popup('Hatalı liste','Geçerli, sıfırdan büyük boylar girin.'); return
            self.boylar=vals; self._yedekle(); p.dismiss(); self.giris_goster()
        box.add_widget(self.btn('LİSTEYİ KAYDET',save)); p.open()

    def excele(self):
        if Workbook is None: self.popup('Eksik paket','Excel aktarımı için openpyxl paketi APK içine eklenmelidir.'); return
        if not self.kayitlar: self.popup('Kayıt yok','Önce kayıt ekleyin.'); return
        try:
            # Android uygulamasının özel klasörüne yazar; dosya uygulama klasöründe oluşur.
            path=os.path.join(self.user_data_dir, 'Orman_Barkod_Kayit_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.xlsx')
            wb=Workbook(); ws=wb.active; ws.title='Barkod Kayıtları'
            headers=['No','Barkod','Ağaç cinsi','Ürün','Boy (m)','Çap (cm)','Adet','Birim hacim (m³)','Toplam hacim (m³)','Açıklama','Kayıt zamanı']; ws.append(headers)
            for i,r in enumerate(self.kayitlar,1): ws.append([i]+[r[k] for k in headers[1:]])
            self.format_ws(ws)
            for row in range(2,ws.max_row+1): ws.cell(row,8).number_format='0.000'; ws.cell(row,9).number_format='0.000'
            sm=wb.create_sheet('Özet'); sm.append(['AĞAÇ CİNSİ','ÜRÜN','KAYIT SAYISI','TOPLAM ADET','TOPLAM HACİM (m³)']); gr={}
            for r in self.kayitlar:
                key=(r['Ağaç cinsi'],r['Ürün']); g=gr.setdefault(key,[0,0,0.0]); g[0]+=1; g[1]+=r['Adet']; g[2]+=r['Toplam hacim (m³)']
            for (c,u),(n,a,v) in sorted(gr.items()): sm.append([c,u,n,a,v])
            sm.append(['GENEL TOPLAM','',len(self.kayitlar),sum(r['Adet'] for r in self.kayitlar),sum(r['Toplam hacim (m³)'] for r in self.kayitlar)]); self.format_ws(sm)
            for row in range(2,sm.max_row+1): sm.cell(row,5).number_format='0.000'
            wb.save(path); self.popup('Excel oluşturuldu',f'Dosya uygulamanın özel klasörüne kaydedildi:\n{path}\n\nAndroid sürümünde paylaşma/İndirilenler klasörüne kaydetme ayrıca entegre edilmelidir.')
        except Exception as e: self.popup('Excel hatası',str(e))

    def format_ws(self,ws):
        ws.freeze_panes='A2'; ws.auto_filter.ref=ws.dimensions
        for c in ws[1]: c.font=Font(bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='16333B'); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        for j in range(1,ws.max_column+1): ws.column_dimensions[get_column_letter(j)].width=min(max(max(len(str(ws.cell(i,j).value or '')) for i in range(1,ws.max_row+1))+2,12),28)

    def _yedekle(self):
        try: Path(self.veri_dosyasi).write_text(json.dumps({'kayitlar':self.kayitlar,'boylar':self.boylar},ensure_ascii=False,indent=2),encoding='utf-8')
        except Exception: pass
    def _yukle(self):
        try:
            if os.path.exists(self.veri_dosyasi):
                d=json.loads(Path(self.veri_dosyasi).read_text(encoding='utf-8')); self.kayitlar=d.get('kayitlar',[]); self.boylar=d.get('boylar',self.boylar)
        except Exception: pass

if __name__ == '__main__':
    Builder.load_string(KV)
    OrmanBarkodAndroid().run()
