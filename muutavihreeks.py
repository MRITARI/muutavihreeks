from textual.app import App, ComposeResult
from textual.screen import Screen
from textual.containers import Horizontal, Vertical, Middle, Center
from textual.widgets import Input, Label, RichLog, RadioSet, RadioButton
import xlwings as xw
from datetime import datetime
from rich.markup import escape

class SplashScreen(Screen): 

    #splash screen olevinaan vähä
    CSS = """
    SplashScreen {
        align: center middle;
        background: #1e1e1e;
    }
    #title-box {
        content-align: center middle;
        text-align: center;
    }
    #big-title {
        text-style: bold;
        color: #92d050;
        margin-bottom: 1;
    }
    #sub-title {
        color: #888888;
    }
    """
    
    def compose(self) -> ComposeResult:
        with Middle():
            with Center(id="title-box"):
                yield Label("[bold #92d050]muutavihreeks.exe[/bold #92d050]", id="big-title")
                yield Label("Rakkaudella MRITARI & Google Gemini", id="sub-title")
    #2 sekunttia
    def on_mount(self) -> None:
                       #
        self.set_timer(2, self.siirry_paanakymaan)
    
    def siirry_paanakymaan(self) -> None:
        self.app.switch_screen(PaaNakyma())

class PaaNakyma(Screen):
    #Päänäkymä
    CSS = """
    PaaNakyma {
        layout: horizontal;
    }
    #left-pane {
        width: 30%;
        height: 100%;
        border: solid #00aaff;
        padding: 1 2;
        background: #1e1e1e;
    }
    #right-pane {
        width: 70%;
        height: 100%;
        border: solid #444444;
        background: #121212;
    }
    #title {
        text-style: bold;
        color: white;
        margin-bottom: 2;
    }
    #status {
        margin-top: 2;
        color: yellow;
    }
    Input {
        margin-top: 1;
    }
    RadioSet {
        margin-top: 2;
        background: #1e1e1e;
    }
    """

    def compose(self) -> ComposeResult:
        with Horizontal():
            # Vasen paneeli
            with Vertical(id="left-pane"):
                yield Label("MUUTAVIHIRIÄKSI", id="title")
                yield Label("Odotetaan skannausta...", id="subtitle")
                yield Input(placeholder="Lue koodi tähän...", id="scan_input")
                yield Label("Tila: Alustetaan...", id="status")
                yield RadioSet(
                    RadioButton("Vihreä", id="c_vihrea", value=True),
                    RadioButton("Sininen", id="c_sininen"),
                    RadioButton("Keltainen", id="c_keltainen"),
                    RadioButton("Oranssi", id="c_oranssi"),
                    RadioButton("Pinkki", id="c_pinkki"),
                    RadioButton("Harmaa", id="c_harmaa"),
                    RadioButton("Vaaleansininen", id="c_vaaleansininen"),
                    id="color_selector"
                )
            # Oikea paneeli
            with Vertical(id="right-pane"):
                yield RichLog(id="log", highlight=True, markup=True)


    #muuttujat ja yrittää yhdistää exceliin
    def on_mount(self) -> None:
        self.sheet = None
        self.aktiivinen_vari = (146, 208, 80)
        self.aktiivinen_vari_nimi = "Vihreä"
        self.vari_nimet = {
            (146, 208, 80): "Vihreä",
            (0, 176, 240): "Sininen",
            (255, 255, 0): "Keltainen",
            (244, 176, 132): "Oranssi",
            (218, 112, 214): "Pinkki",
            (191, 191, 191): "Harmaa",
            (204, 238, 255): "Vaaleansininen"
        }
        self.log_widget = self.query_one(RichLog)
        self.status_label = self.query_one("#status", Label)
        self.scan_input = self.query_one(Input)
        self.scan_input.focus()
        self.yhdista_exceliin()

    def hae_varin_nimi(self, rgb_tuple) -> str:
        if rgb_tuple is None or rgb_tuple == (255, 255, 255):
            return "Ei väriä"
        return self.vari_nimet.get(rgb_tuple, f"Muu väri {rgb_tuple}")

    def on_radio_set_changed(self, event: RadioSet.Changed) -> None:
        varit = {
            "c_vihrea": (146, 208, 80),
            "c_sininen": (0, 176, 240),
            "c_keltainen": (255, 255, 0),
            "c_oranssi": (244, 176, 132),
            "c_pinkki": (218, 112, 214),
            "c_harmaa": (191, 191, 191),
            "c_vaaleansininen": (204, 238, 255)
        }
        self.aktiivinen_vari = varit.get(event.pressed.id, (146, 208, 80))
        self.aktiivinen_vari_nimi = str(event.pressed.label)
        self.scan_input.focus()

    #Käyttää xlwings-kirjastoa yhdistääkseen auki olevaan excel-ikkunaan
    #Ei saa olla suojatussa näkymässä

    def yhdista_exceliin(self) -> None:
        try:
            wb = xw.books.active
            self.sheet = wb.sheets.active
            self.loki_viesti(f"[bold cyan]muutavihreeks: Yhdistetty tiedostoon {wb.name}[/bold cyan]")
            self.status_label.update("Tila: [bold green]Valmis skannaukseen[/bold green]")
        except Exception:
            self.loki_viesti("[bold red]muutavihreeks: Ei saatu yhteyttä Exceliin. Onko tiedosto auki?[/bold red]")
            self.status_label.update("Tila: [bold red]Ei yhteyttä[/bold red]")

    def loki_viesti(self, viesti: str) -> None:
        aika = datetime.now().strftime("%H:%M:%S")
        self.log_widget.write(f"[[dim]{aika}[/dim]] {viesti}")

    def on_input_submitted(self, event: Input.Submitted) -> None:
        luettu_koodi = event.value.strip().upper()
        self.scan_input.value = ""

        if not luettu_koodi:
            return

        turvallinen_koodi = escape(luettu_koodi)

        if not self.sheet:
            self.loki_viesti("[bold red]Virhe: Ei yhteyttä Exceliin.[/bold red]")
            return

        try:
            loytyi = False
            c_sarakkeen_arvot = self.sheet.range('C1:C10000').value

            for indeksi, solun_arvo in enumerate(c_sarakkeen_arvot):
                if solun_arvo is not None and str(solun_arvo).strip().upper() == luettu_koodi:
                    rivi_nro = indeksi + 1
                    kohde_alue = self.sheet.range(f'{rivi_nro}:{rivi_nro}')
                    
                    nykyinen_vari_rgb = kohde_alue.color
                    nykyinen_vari_nimi = self.hae_varin_nimi(nykyinen_vari_rgb)

                    if nykyinen_vari_rgb == self.aktiivinen_vari:
                        self.loki_viesti(f"[bold yellow]OHITETTU:[/bold yellow] {turvallinen_koodi} (Rivi {rivi_nro}) | Oli jo {nykyinen_vari_nimi}")
                        self.status_label.update(f"Viimeisin: {turvallinen_koodi}\n[bold yellow]Ohitettu (jo {self.aktiivinen_vari_nimi})[/bold yellow]")
                    else:
                        kohde_alue.color = self.aktiivinen_vari
                        self.loki_viesti(f"[bold green]MUUTETTU:[/bold green] {turvallinen_koodi} (Rivi {rivi_nro}) | {nykyinen_vari_nimi} -> {self.aktiivinen_vari_nimi}")
                        self.status_label.update(f"Viimeisin: {turvallinen_koodi}\n[bold green]Muutettu: {self.aktiivinen_vari_nimi}[/bold green]")
                    
                    loytyi = True
                    break

            if not loytyi:
                self.loki_viesti(f"[bold red]EI LÖYTYNYT:[/bold red] {turvallinen_koodi}")
                self.status_label.update(f"Viimeisin: {turvallinen_koodi}\n[bold red]Koodia ei löytynyt[/bold red]")

        except Exception as e:
            turvallinen_virhe = escape(str(e))
            self.loki_viesti(f"[bold red]Excel-virhe: {turvallinen_virhe}[/bold red]")

class MuutavihreeksApp(App):
    TITLE = "muutavihreeks"

    def on_mount(self) -> None:
        self.push_screen(SplashScreen())

if __name__ == "__main__":
    app = MuutavihreeksApp()
    app.run()