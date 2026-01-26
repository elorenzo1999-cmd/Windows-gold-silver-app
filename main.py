import sys
from dataclasses import dataclass
from datetime import datetime, timedelta

import pandas as pd
import pyqtgraph as pg
import yfinance as yf
from PySide6 import QtCore, QtGui, QtWidgets


@dataclass
class AssetQuote:
    name: str
    symbol: str
    price: float | None
    change: float | None
    updated: datetime | None


ASSETS = [
    ("Gold", "GC=F"),
    ("Silver", "SI=F"),
    ("GDXJ", "GDXJ"),
]


class PriceCard(QtWidgets.QFrame):
    def __init__(self, title: str, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("priceCard")
        self.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.setProperty("class", "card")

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(6)

        self.title_label = QtWidgets.QLabel(title)
        self.title_label.setObjectName("cardTitle")
        self.price_label = QtWidgets.QLabel("--")
        self.price_label.setObjectName("cardPrice")
        self.change_label = QtWidgets.QLabel("--")
        self.change_label.setObjectName("cardChange")

        layout.addWidget(self.title_label)
        layout.addStretch()
        layout.addWidget(self.price_label)
        layout.addWidget(self.change_label)

    def update_quote(self, quote: AssetQuote) -> None:
        if quote.price is None:
            self.price_label.setText("--")
            self.change_label.setText("No data")
            self.change_label.setProperty("state", "neutral")
            self.change_label.style().polish(self.change_label)
            return

        self.price_label.setText(f"${quote.price:,.2f}")
        if quote.change is None:
            self.change_label.setText("--")
            self.change_label.setProperty("state", "neutral")
        else:
            sign = "+" if quote.change >= 0 else ""
            self.change_label.setText(f"{sign}{quote.change:,.2f}%")
            self.change_label.setProperty("state", "up" if quote.change >= 0 else "down")
        self.change_label.style().polish(self.change_label)


class GoldSilverApp(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Gold, Silver & GDXJ Dashboard")
        self.resize(1100, 700)

        self.cards: dict[str, PriceCard] = {}
        self.current_symbol = ASSETS[0][1]

        central = QtWidgets.QWidget()
        root_layout = QtWidgets.QVBoxLayout(central)
        root_layout.setContentsMargins(24, 24, 24, 24)
        root_layout.setSpacing(20)

        header_layout = QtWidgets.QHBoxLayout()
        title = QtWidgets.QLabel("Metals & Miners Snapshot")
        title.setObjectName("appTitle")
        subtitle = QtWidgets.QLabel("Live snapshot and interactive chart")
        subtitle.setObjectName("appSubtitle")
        header_text = QtWidgets.QVBoxLayout()
        header_text.addWidget(title)
        header_text.addWidget(subtitle)
        header_layout.addLayout(header_text)
        header_layout.addStretch()

        self.last_updated_label = QtWidgets.QLabel("Last updated: --")
        self.last_updated_label.setObjectName("lastUpdated")
        header_layout.addWidget(self.last_updated_label)
        root_layout.addLayout(header_layout)

        cards_layout = QtWidgets.QHBoxLayout()
        cards_layout.setSpacing(16)
        for name, symbol in ASSETS:
            card = PriceCard(name)
            self.cards[symbol] = card
            card.mousePressEvent = self._make_card_handler(symbol)
            cards_layout.addWidget(card)
        root_layout.addLayout(cards_layout)

        chart_container = QtWidgets.QFrame()
        chart_container.setObjectName("chartContainer")
        chart_layout = QtWidgets.QVBoxLayout(chart_container)
        chart_layout.setContentsMargins(12, 12, 12, 12)
        chart_layout.setSpacing(8)

        chart_header = QtWidgets.QHBoxLayout()
        self.chart_title = QtWidgets.QLabel("Gold - 6 Month Trend")
        self.chart_title.setObjectName("chartTitle")
        chart_header.addWidget(self.chart_title)
        chart_header.addStretch()

        self.range_combo = QtWidgets.QComboBox()
        self.range_combo.addItems(["1M", "3M", "6M", "1Y", "2Y"])
        self.range_combo.setCurrentText("6M")
        self.range_combo.currentTextChanged.connect(self.refresh_chart)
        chart_header.addWidget(self.range_combo)

        chart_layout.addLayout(chart_header)

        self.chart = pg.PlotWidget()
        self.chart.setBackground(None)
        self.chart.showGrid(x=True, y=True, alpha=0.2)
        self.chart.addLegend()
        self.chart_layout_item = self.chart.getPlotItem()
        self.chart_layout_item.setLabel("left", "Price", units="USD")
        self.chart_layout_item.setLabel("bottom", "Date")
        self.chart_layout_item.setMenuEnabled(False)
        self.chart_layout_item.hideButtons()
        chart_layout.addWidget(self.chart)

        root_layout.addWidget(chart_container)

        self.setCentralWidget(central)

        self._apply_theme()

        self.refresh_data()
        self.refresh_chart()

        self.timer = QtCore.QTimer(self)
        self.timer.setInterval(60_000)
        self.timer.timeout.connect(self.refresh_data)
        self.timer.start()

    def _apply_theme(self) -> None:
        QtWidgets.QApplication.setStyle("Fusion")
        palette = QtGui.QPalette()
        palette.setColor(QtGui.QPalette.Window, QtGui.QColor("#0f172a"))
        palette.setColor(QtGui.QPalette.WindowText, QtGui.QColor("#e2e8f0"))
        palette.setColor(QtGui.QPalette.Base, QtGui.QColor("#0f172a"))
        palette.setColor(QtGui.QPalette.AlternateBase, QtGui.QColor("#1e293b"))
        palette.setColor(QtGui.QPalette.Text, QtGui.QColor("#e2e8f0"))
        palette.setColor(QtGui.QPalette.Button, QtGui.QColor("#1e293b"))
        palette.setColor(QtGui.QPalette.ButtonText, QtGui.QColor("#e2e8f0"))
        palette.setColor(QtGui.QPalette.Highlight, QtGui.QColor("#38bdf8"))
        palette.setColor(QtGui.QPalette.HighlightedText, QtGui.QColor("#0f172a"))
        self.setPalette(palette)

        self.setStyleSheet(
            """
            QLabel#appTitle { font-size: 26px; font-weight: 600; }
            QLabel#appSubtitle { color: #94a3b8; }
            QLabel#lastUpdated { color: #94a3b8; }
            QFrame#priceCard { background: #1e293b; border-radius: 16px; }
            QLabel#cardTitle { color: #94a3b8; font-size: 14px; }
            QLabel#cardPrice { font-size: 26px; font-weight: 600; }
            QLabel#cardChange[state="up"] { color: #22c55e; font-weight: 600; }
            QLabel#cardChange[state="down"] { color: #ef4444; font-weight: 600; }
            QLabel#cardChange[state="neutral"] { color: #94a3b8; }
            QFrame#chartContainer { background: #0b1220; border-radius: 20px; }
            QLabel#chartTitle { font-size: 18px; font-weight: 600; }
            QComboBox { padding: 6px 10px; border-radius: 10px; background: #1e293b; }
            """
        )

    def _make_card_handler(self, symbol: str):
        def handler(event):
            self.current_symbol = symbol
            self.refresh_chart()

        return handler

    def refresh_data(self) -> None:
        quotes = self._fetch_quotes()
        for quote in quotes:
            card = self.cards.get(quote.symbol)
            if card:
                card.update_quote(quote)

        last = max((q.updated for q in quotes if q.updated), default=None)
        if last:
            self.last_updated_label.setText(f"Last updated: {last.strftime('%Y-%m-%d %H:%M:%S')}")
        else:
            self.last_updated_label.setText("Last updated: --")

    def refresh_chart(self) -> None:
        self.chart.clear()
        symbol = self.current_symbol
        name = next(name for name, sym in ASSETS if sym == symbol)
        range_text = self.range_combo.currentText()
        self.chart_title.setText(f"{name} - {range_text} Trend")

        history = self._fetch_history(symbol, range_text)
        if history.empty:
            self.chart.plot([], [])
            return

        dates = history.index.to_pydatetime()
        prices = history["Close"].values
        pen = pg.mkPen(color="#38bdf8", width=2)
        self.chart.plot(dates, prices, pen=pen, name=f"{name} Close")
        self.chart_layout_item.setLimits(xMin=min(dates).timestamp(), xMax=max(dates).timestamp())

    def _fetch_quotes(self) -> list[AssetQuote]:
        results = []
        for name, symbol in ASSETS:
            ticker = yf.Ticker(symbol)
            info = ticker.fast_info
            price = info.get("lastPrice") if info else None
            previous = info.get("previousClose") if info else None
            change = None
            if price is not None and previous:
                change = ((price - previous) / previous) * 100
            results.append(AssetQuote(name=name, symbol=symbol, price=price, change=change, updated=datetime.now()))
        return results

    def _fetch_history(self, symbol: str, range_text: str) -> pd.DataFrame:
        ranges = {
            "1M": timedelta(days=30),
            "3M": timedelta(days=90),
            "6M": timedelta(days=180),
            "1Y": timedelta(days=365),
            "2Y": timedelta(days=730),
        }
        end = datetime.now()
        start = end - ranges.get(range_text, timedelta(days=180))
        data = yf.download(symbol, start=start, end=end, progress=False)
        return data


def main() -> None:
    app = QtWidgets.QApplication(sys.argv)
    window = GoldSilverApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
