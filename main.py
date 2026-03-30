"""Microsoft 365 Environment Manager — desktop GUI."""

import subprocess
import sys


def _bootstrap():
    """Auto-install any missing dependencies before the app starts."""
    pkgs = {"PySide6": "PySide6", "msal": "msal", "requests": "requests"}
    missing = []
    for import_name, install_name in pkgs.items():
        try:
            __import__(import_name.lower())
        except ImportError:
            missing.append(install_name)
    if missing:
        print(f"Installing missing packages: {', '.join(missing)} ...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--quiet"] + missing
        )
        print("Done. Launching app...")


_bootstrap()

import json
import threading

from PySide6 import QtCore, QtGui, QtWidgets

from auth import AuthError, M365Auth
from graph_api import GraphAPI, GraphError


# ── Dark theme palette ────────────────────────────────────────────────────────

def _apply_theme(app: QtWidgets.QApplication) -> None:
    app.setStyle("Fusion")
    pal = QtGui.QPalette()
    pal.setColor(QtGui.QPalette.Window, QtGui.QColor("#0f172a"))
    pal.setColor(QtGui.QPalette.WindowText, QtGui.QColor("#e2e8f0"))
    pal.setColor(QtGui.QPalette.Base, QtGui.QColor("#1e293b"))
    pal.setColor(QtGui.QPalette.AlternateBase, QtGui.QColor("#0f172a"))
    pal.setColor(QtGui.QPalette.Text, QtGui.QColor("#e2e8f0"))
    pal.setColor(QtGui.QPalette.Button, QtGui.QColor("#1e293b"))
    pal.setColor(QtGui.QPalette.ButtonText, QtGui.QColor("#e2e8f0"))
    pal.setColor(QtGui.QPalette.Highlight, QtGui.QColor("#38bdf8"))
    pal.setColor(QtGui.QPalette.HighlightedText, QtGui.QColor("#0f172a"))
    pal.setColor(QtGui.QPalette.ToolTipBase, QtGui.QColor("#1e293b"))
    pal.setColor(QtGui.QPalette.ToolTipText, QtGui.QColor("#e2e8f0"))
    app.setPalette(pal)
    app.setStyleSheet(
        """
        QWidget { font-size: 13px; }
        QMainWindow, QDialog { background: #0f172a; }
        QTabWidget::pane { border: 1px solid #334155; border-radius: 8px; }
        QTabBar::tab { background: #1e293b; color: #94a3b8; padding: 8px 20px;
                       border-top-left-radius: 6px; border-top-right-radius: 6px; }
        QTabBar::tab:selected { background: #0f172a; color: #e2e8f0; }
        QTableWidget { gridline-color: #1e293b; border: none; }
        QTableWidget::item:selected { background: #38bdf8; color: #0f172a; }
        QHeaderView::section { background: #1e293b; color: #94a3b8;
                               padding: 6px; border: none; }
        QPushButton { background: #1e293b; color: #e2e8f0; border-radius: 6px;
                      padding: 6px 16px; border: 1px solid #334155; }
        QPushButton:hover { background: #334155; }
        QPushButton#primary { background: #38bdf8; color: #0f172a;
                               border: none; font-weight: 600; }
        QPushButton#primary:hover { background: #7dd3fc; }
        QPushButton#danger { background: #ef4444; color: #fff; border: none; }
        QPushButton#danger:hover { background: #f87171; }
        QLineEdit, QComboBox, QTextEdit, QPlainTextEdit {
            background: #1e293b; color: #e2e8f0; border: 1px solid #334155;
            border-radius: 6px; padding: 6px 10px; }
        QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus {
            border-color: #38bdf8; }
        QLabel#heading { font-size: 18px; font-weight: 600; }
        QLabel#subheading { color: #94a3b8; }
        QLabel#status_ok { color: #22c55e; }
        QLabel#status_err { color: #ef4444; }
        QFrame#card { background: #1e293b; border-radius: 10px; }
        QScrollBar:vertical { background: #1e293b; width: 8px; border-radius: 4px; }
        QScrollBar::handle:vertical { background: #334155; border-radius: 4px; }
        """
    )


# ── Login widget ──────────────────────────────────────────────────────────────

class LoginWidget(QtWidgets.QWidget):
    logged_in = QtCore.Signal(M365Auth, GraphAPI)

    def __init__(self, parent=None):
        super().__init__(parent)
        outer = QtWidgets.QVBoxLayout(self)
        outer.setAlignment(QtCore.Qt.AlignCenter)

        card = QtWidgets.QFrame()
        card.setObjectName("card")
        card.setFixedWidth(420)
        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(32, 32, 32, 32)
        layout.setSpacing(16)

        title = QtWidgets.QLabel("Microsoft 365 Manager")
        title.setObjectName("heading")
        title.setAlignment(QtCore.Qt.AlignCenter)

        sub = QtWidgets.QLabel("Sign in with your M365 admin account")
        sub.setObjectName("subheading")
        sub.setAlignment(QtCore.Qt.AlignCenter)

        hint_label = QtWidgets.QLabel("Username (optional — pre-fills the browser login)")
        hint_label.setObjectName("subheading")

        self.username_input = QtWidgets.QLineEdit()
        self.username_input.setPlaceholderText("admin@yourtenant.onmicrosoft.com")
        self.username_input.returnPressed.connect(self._on_connect)

        self.connect_btn = QtWidgets.QPushButton("Sign In  →  Opens browser with MFA")
        self.connect_btn.setObjectName("primary")
        self.connect_btn.setFixedHeight(44)
        self.connect_btn.clicked.connect(self._on_connect)

        self.status_label = QtWidgets.QLabel("")
        self.status_label.setAlignment(QtCore.Qt.AlignCenter)
        self.status_label.setWordWrap(True)

        note = QtWidgets.QLabel(
            "A Microsoft login page will open in your browser.\n"
            "Complete sign-in and MFA there — the app will connect automatically."
        )
        note.setObjectName("subheading")
        note.setAlignment(QtCore.Qt.AlignCenter)
        note.setWordWrap(True)

        layout.addWidget(title)
        layout.addWidget(sub)
        layout.addSpacing(8)
        layout.addWidget(hint_label)
        layout.addWidget(self.username_input)
        layout.addSpacing(4)
        layout.addWidget(self.connect_btn)
        layout.addWidget(self.status_label)
        layout.addSpacing(4)
        layout.addWidget(note)

        outer.addWidget(card)

        self._auth = M365Auth()

    def _on_connect(self):
        login_hint = self.username_input.text().strip()

        self.connect_btn.setEnabled(False)
        self.connect_btn.setText("Waiting for browser sign-in…")
        self.status_label.setObjectName("")
        self.status_label.setText("")

        def _worker():
            try:
                token = self._auth.login_interactive(login_hint=login_hint)
                api = GraphAPI(token)
                QtCore.QMetaObject.invokeMethod(
                    self, "_on_success",
                    QtCore.Qt.QueuedConnection,
                    QtCore.Q_ARG(object, api),
                )
            except (AuthError, Exception) as exc:
                QtCore.QMetaObject.invokeMethod(
                    self, "_on_failure",
                    QtCore.Qt.QueuedConnection,
                    QtCore.Q_ARG(str, str(exc)),
                )

        threading.Thread(target=_worker, daemon=True).start()

    @QtCore.Slot(object)
    def _on_success(self, api: GraphAPI):
        self.connect_btn.setEnabled(True)
        self.connect_btn.setText("Sign In  →  Opens browser with MFA")
        self.logged_in.emit(self._auth, api)

    @QtCore.Slot(str)
    def _on_failure(self, message: str):
        self.connect_btn.setEnabled(True)
        self.connect_btn.setText("Sign In  →  Opens browser with MFA")
        self._set_error(message)

    def _set_error(self, msg: str):
        self.status_label.setObjectName("status_err")
        self.status_label.setText(msg)
        self.status_label.style().polish(self.status_label)


# ── User dialog ───────────────────────────────────────────────────────────────

class UserDialog(QtWidgets.QDialog):
    """Create or edit a user."""

    def __init__(self, api: GraphAPI, user: dict | None = None, parent=None):
        super().__init__(parent)
        self._api = api
        self._user = user
        self.setWindowTitle("New User" if user is None else "Edit User")
        self.setMinimumWidth(440)

        layout = QtWidgets.QFormLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(12)

        self.display_name = QtWidgets.QLineEdit(user.get("displayName", "") if user else "")
        self.upn = QtWidgets.QLineEdit(user.get("userPrincipalName", "") if user else "")
        self.job_title = QtWidgets.QLineEdit(user.get("jobTitle", "") if user else "")
        self.department = QtWidgets.QLineEdit(user.get("department", "") if user else "")

        self.password = QtWidgets.QLineEdit()
        self.password.setEchoMode(QtWidgets.QLineEdit.Password)
        self.password.setPlaceholderText("Leave blank to keep existing" if user else "Required")

        self.enabled = QtWidgets.QCheckBox("Account enabled")
        self.enabled.setChecked(user.get("accountEnabled", True) if user else True)

        layout.addRow("Display Name *", self.display_name)
        layout.addRow("User Principal Name *", self.upn)
        layout.addRow("Job Title", self.job_title)
        layout.addRow("Department", self.department)
        layout.addRow("Password" + ("" if user else " *"), self.password)
        layout.addRow("", self.enabled)

        self.status_label = QtWidgets.QLabel("")
        self.status_label.setWordWrap(True)
        layout.addRow(self.status_label)

        btns = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel
        )
        btns.accepted.connect(self._save)
        btns.rejected.connect(self.reject)
        layout.addRow(btns)

    def _save(self):
        display_name = self.display_name.text().strip()
        upn = self.upn.text().strip()
        password = self.password.text()

        if not display_name or not upn:
            self.status_label.setObjectName("status_err")
            self.status_label.setText("Display Name and UPN are required.")
            self.status_label.style().polish(self.status_label)
            return

        if self._user is None and not password:
            self.status_label.setObjectName("status_err")
            self.status_label.setText("Password is required for new users.")
            self.status_label.style().polish(self.status_label)
            return

        try:
            if self._user is None:
                payload = {
                    "displayName": display_name,
                    "userPrincipalName": upn,
                    "accountEnabled": self.enabled.isChecked(),
                    "passwordProfile": {
                        "forceChangePasswordNextSignIn": False,
                        "password": password,
                    },
                }
                if self.job_title.text().strip():
                    payload["jobTitle"] = self.job_title.text().strip()
                if self.department.text().strip():
                    payload["department"] = self.department.text().strip()
                self._api.create_user(payload)
            else:
                payload = {
                    "displayName": display_name,
                    "userPrincipalName": upn,
                    "accountEnabled": self.enabled.isChecked(),
                    "jobTitle": self.job_title.text().strip(),
                    "department": self.department.text().strip(),
                }
                if password:
                    payload["passwordProfile"] = {
                        "forceChangePasswordNextSignIn": False,
                        "password": password,
                    }
                self._api.update_user(self._user["id"], payload)
            self.accept()
        except (GraphError, Exception) as exc:
            self.status_label.setObjectName("status_err")
            self.status_label.setText(str(exc))
            self.status_label.style().polish(self.status_label)


# ── License dialog ────────────────────────────────────────────────────────────

class ManageLicensesDialog(QtWidgets.QDialog):
    def __init__(self, api: GraphAPI, user: dict, skus: list[dict], parent=None):
        super().__init__(parent)
        self._api = api
        self._user = user
        self._skus = skus
        self.setWindowTitle(f"Licenses — {user.get('displayName', user['id'])}")
        self.setMinimumSize(520, 400)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        layout.addWidget(QtWidgets.QLabel("Available license SKUs:"))

        self.sku_table = QtWidgets.QTableWidget(0, 4)
        self.sku_table.setHorizontalHeaderLabels(["SKU Name", "Total", "Used", "Assigned"])
        self.sku_table.horizontalHeader().setStretchLastSection(False)
        self.sku_table.horizontalHeader().setSectionResizeMode(
            0, QtWidgets.QHeaderView.Stretch
        )
        self.sku_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.sku_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        layout.addWidget(self.sku_table)

        btn_row = QtWidgets.QHBoxLayout()
        assign_btn = QtWidgets.QPushButton("Assign Selected")
        assign_btn.setObjectName("primary")
        assign_btn.clicked.connect(self._assign)
        remove_btn = QtWidgets.QPushButton("Remove Selected")
        remove_btn.setObjectName("danger")
        remove_btn.clicked.connect(self._remove)
        btn_row.addWidget(assign_btn)
        btn_row.addWidget(remove_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        self.status_label = QtWidgets.QLabel("")
        layout.addWidget(self.status_label)

        close_btn = QtWidgets.QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)

        self._load()

    def _load(self):
        assigned_ids = {
            lic["skuId"]
            for lic in self._api.get_user_licenses(self._user["id"])
        }
        self.sku_table.setRowCount(0)
        self._sku_ids = []
        for sku in self._skus:
            row = self.sku_table.rowCount()
            self.sku_table.insertRow(row)
            name = sku.get("skuPartNumber", sku["skuId"])
            total = sku.get("prepaidUnits", {}).get("enabled", 0)
            used = sku.get("consumedUnits", 0)
            is_assigned = sku["skuId"] in assigned_ids
            self.sku_table.setItem(row, 0, QtWidgets.QTableWidgetItem(name))
            self.sku_table.setItem(row, 1, QtWidgets.QTableWidgetItem(str(total)))
            self.sku_table.setItem(row, 2, QtWidgets.QTableWidgetItem(str(used)))
            assigned_item = QtWidgets.QTableWidgetItem("Yes" if is_assigned else "No")
            if is_assigned:
                assigned_item.setForeground(QtGui.QColor("#22c55e"))
            self.sku_table.setItem(row, 3, assigned_item)
            self._sku_ids.append(sku["skuId"])

    def _selected_sku_id(self):
        rows = self.sku_table.selectedItems()
        if not rows:
            return None
        return self._sku_ids[self.sku_table.currentRow()]

    def _assign(self):
        sku_id = self._selected_sku_id()
        if not sku_id:
            return
        try:
            self._api.assign_license(self._user["id"], sku_id)
            self._load()
            self.status_label.setObjectName("status_ok")
            self.status_label.setText("License assigned.")
            self.status_label.style().polish(self.status_label)
        except (GraphError, Exception) as exc:
            self.status_label.setObjectName("status_err")
            self.status_label.setText(str(exc))
            self.status_label.style().polish(self.status_label)

    def _remove(self):
        sku_id = self._selected_sku_id()
        if not sku_id:
            return
        try:
            self._api.remove_license(self._user["id"], sku_id)
            self._load()
            self.status_label.setObjectName("status_ok")
            self.status_label.setText("License removed.")
            self.status_label.style().polish(self.status_label)
        except (GraphError, Exception) as exc:
            self.status_label.setObjectName("status_err")
            self.status_label.setText(str(exc))
            self.status_label.style().polish(self.status_label)


# ── Users tab ─────────────────────────────────────────────────────────────────

class UsersTab(QtWidgets.QWidget):
    def __init__(self, api: GraphAPI, parent=None):
        super().__init__(parent)
        self._api = api
        self._users: list[dict] = []
        self._skus: list[dict] = []

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # Toolbar
        toolbar = QtWidgets.QHBoxLayout()
        self.search_input = QtWidgets.QLineEdit()
        self.search_input.setPlaceholderText("Search users…")
        self.search_input.textChanged.connect(self._filter)

        refresh_btn = QtWidgets.QPushButton("Refresh")
        refresh_btn.clicked.connect(self._load)

        new_btn = QtWidgets.QPushButton("New User")
        new_btn.setObjectName("primary")
        new_btn.clicked.connect(self._new_user)

        self.edit_btn = QtWidgets.QPushButton("Edit")
        self.edit_btn.clicked.connect(self._edit_user)

        self.delete_btn = QtWidgets.QPushButton("Delete")
        self.delete_btn.setObjectName("danger")
        self.delete_btn.clicked.connect(self._delete_user)

        self.license_btn = QtWidgets.QPushButton("Manage Licenses")
        self.license_btn.clicked.connect(self._manage_licenses)

        toolbar.addWidget(self.search_input)
        toolbar.addWidget(refresh_btn)
        toolbar.addStretch()
        toolbar.addWidget(new_btn)
        toolbar.addWidget(self.edit_btn)
        toolbar.addWidget(self.delete_btn)
        toolbar.addWidget(self.license_btn)
        layout.addLayout(toolbar)

        # Table
        self.table = QtWidgets.QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(
            ["Display Name", "Email / UPN", "Job Title", "Department", "Enabled"]
        )
        self.table.horizontalHeader().setSectionResizeMode(
            1, QtWidgets.QHeaderView.Stretch
        )
        self.table.horizontalHeader().setSectionResizeMode(
            0, QtWidgets.QHeaderView.ResizeToContents
        )
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        layout.addWidget(self.table)

        self.status_label = QtWidgets.QLabel("")
        self.status_label.setObjectName("subheading")
        layout.addWidget(self.status_label)

        self._load()

    def _load(self):
        self.status_label.setText("Loading users…")
        QtWidgets.QApplication.processEvents()

        def _worker():
            try:
                users = self._api.get_users()
                skus = self._api.get_subscribed_skus()
                QtCore.QMetaObject.invokeMethod(
                    self, "_populate",
                    QtCore.Qt.QueuedConnection,
                    QtCore.Q_ARG(object, users),
                    QtCore.Q_ARG(object, skus),
                )
            except (GraphError, Exception) as exc:
                QtCore.QMetaObject.invokeMethod(
                    self, "_set_error",
                    QtCore.Qt.QueuedConnection,
                    QtCore.Q_ARG(str, str(exc)),
                )

        threading.Thread(target=_worker, daemon=True).start()

    @QtCore.Slot(object, object)
    def _populate(self, users: list[dict], skus: list[dict]):
        self._users = users
        self._skus = skus
        self._render(users)
        self.status_label.setText(f"{len(users)} user(s) loaded.")

    @QtCore.Slot(str)
    def _set_error(self, msg: str):
        self.status_label.setObjectName("status_err")
        self.status_label.setText(msg)
        self.status_label.style().polish(self.status_label)

    def _render(self, users: list[dict]):
        self.table.setRowCount(0)
        for user in users:
            row = self.table.rowCount()
            self.table.insertRow(row)
            enabled = user.get("accountEnabled", False)
            items = [
                user.get("displayName", ""),
                user.get("userPrincipalName", ""),
                user.get("jobTitle") or "",
                user.get("department") or "",
                "Yes" if enabled else "No",
            ]
            for col, text in enumerate(items):
                item = QtWidgets.QTableWidgetItem(text)
                if col == 4:
                    item.setForeground(
                        QtGui.QColor("#22c55e") if enabled else QtGui.QColor("#ef4444")
                    )
                self.table.setItem(row, col, item)
            # Store user id in first item
            self.table.item(row, 0).setData(QtCore.Qt.UserRole, user["id"])

    def _filter(self, text: str):
        text = text.lower()
        filtered = [
            u for u in self._users
            if text in (u.get("displayName") or "").lower()
            or text in (u.get("userPrincipalName") or "").lower()
        ]
        self._render(filtered)

    def _selected_user(self) -> dict | None:
        row = self.table.currentRow()
        if row < 0:
            return None
        user_id = self.table.item(row, 0).data(QtCore.Qt.UserRole)
        return next((u for u in self._users if u["id"] == user_id), None)

    def _new_user(self):
        dlg = UserDialog(self._api, parent=self)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            self._load()

    def _edit_user(self):
        user = self._selected_user()
        if not user:
            QtWidgets.QMessageBox.information(self, "No selection", "Select a user first.")
            return
        dlg = UserDialog(self._api, user=user, parent=self)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            self._load()

    def _delete_user(self):
        user = self._selected_user()
        if not user:
            QtWidgets.QMessageBox.information(self, "No selection", "Select a user first.")
            return
        name = user.get("displayName", user["id"])
        reply = QtWidgets.QMessageBox.warning(
            self, "Delete User",
            f"Permanently delete '{name}'? This cannot be undone.",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.Cancel,
        )
        if reply != QtWidgets.QMessageBox.Yes:
            return
        try:
            self._api.delete_user(user["id"])
            self._load()
        except (GraphError, Exception) as exc:
            QtWidgets.QMessageBox.critical(self, "Error", str(exc))

    def _manage_licenses(self):
        user = self._selected_user()
        if not user:
            QtWidgets.QMessageBox.information(self, "No selection", "Select a user first.")
            return
        dlg = ManageLicensesDialog(self._api, user, self._skus, parent=self)
        dlg.exec()


# ── Licenses tab ──────────────────────────────────────────────────────────────

class LicensesTab(QtWidgets.QWidget):
    def __init__(self, api: GraphAPI, parent=None):
        super().__init__(parent)
        self._api = api

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        toolbar = QtWidgets.QHBoxLayout()
        refresh_btn = QtWidgets.QPushButton("Refresh")
        refresh_btn.clicked.connect(self._load)
        toolbar.addStretch()
        toolbar.addWidget(refresh_btn)
        layout.addLayout(toolbar)

        self.table = QtWidgets.QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(
            ["License Name (SKU)", "Total Units", "Consumed", "Remaining"]
        )
        self.table.horizontalHeader().setSectionResizeMode(
            0, QtWidgets.QHeaderView.Stretch
        )
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        layout.addWidget(self.table)

        self.status_label = QtWidgets.QLabel("")
        self.status_label.setObjectName("subheading")
        layout.addWidget(self.status_label)

        self._load()

    def _load(self):
        self.status_label.setText("Loading licenses…")

        def _worker():
            try:
                skus = self._api.get_subscribed_skus()
                QtCore.QMetaObject.invokeMethod(
                    self, "_populate",
                    QtCore.Qt.QueuedConnection,
                    QtCore.Q_ARG(object, skus),
                )
            except (GraphError, Exception) as exc:
                QtCore.QMetaObject.invokeMethod(
                    self, "_set_error",
                    QtCore.Qt.QueuedConnection,
                    QtCore.Q_ARG(str, str(exc)),
                )

        threading.Thread(target=_worker, daemon=True).start()

    @QtCore.Slot(object)
    def _populate(self, skus: list[dict]):
        self.table.setRowCount(0)
        for sku in skus:
            row = self.table.rowCount()
            self.table.insertRow(row)
            name = sku.get("skuPartNumber", sku["skuId"])
            total = sku.get("prepaidUnits", {}).get("enabled", 0)
            consumed = sku.get("consumedUnits", 0)
            remaining = max(0, total - consumed)
            self.table.setItem(row, 0, QtWidgets.QTableWidgetItem(name))
            self.table.setItem(row, 1, QtWidgets.QTableWidgetItem(str(total)))
            self.table.setItem(row, 2, QtWidgets.QTableWidgetItem(str(consumed)))
            rem_item = QtWidgets.QTableWidgetItem(str(remaining))
            rem_item.setForeground(
                QtGui.QColor("#22c55e") if remaining > 0 else QtGui.QColor("#ef4444")
            )
            self.table.setItem(row, 3, rem_item)
        self.status_label.setText(f"{len(skus)} SKU(s) loaded.")

    @QtCore.Slot(str)
    def _set_error(self, msg: str):
        self.status_label.setObjectName("status_err")
        self.status_label.setText(msg)
        self.status_label.style().polish(self.status_label)


# ── Graph Explorer tab ────────────────────────────────────────────────────────

class GraphExplorerTab(QtWidgets.QWidget):
    def __init__(self, api: GraphAPI, parent=None):
        super().__init__(parent)
        self._api = api

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # Request bar
        request_row = QtWidgets.QHBoxLayout()
        self.method_combo = QtWidgets.QComboBox()
        self.method_combo.addItems(["GET", "POST", "PATCH", "DELETE"])
        self.method_combo.setFixedWidth(90)
        self.method_combo.currentTextChanged.connect(self._toggle_body)

        self.endpoint_input = QtWidgets.QLineEdit()
        self.endpoint_input.setPlaceholderText("/users  or  /me  or  /groups/…")
        self.endpoint_input.returnPressed.connect(self._run)

        run_btn = QtWidgets.QPushButton("Run")
        run_btn.setObjectName("primary")
        run_btn.setFixedWidth(70)
        run_btn.clicked.connect(self._run)

        request_row.addWidget(self.method_combo)
        request_row.addWidget(self.endpoint_input)
        request_row.addWidget(run_btn)
        layout.addLayout(request_row)

        # Body
        self.body_label = QtWidgets.QLabel("Request Body (JSON):")
        self.body_input = QtWidgets.QPlainTextEdit()
        self.body_input.setPlaceholderText('{ "key": "value" }')
        self.body_input.setFixedHeight(100)
        layout.addWidget(self.body_label)
        layout.addWidget(self.body_input)

        # Response
        layout.addWidget(QtWidgets.QLabel("Response:"))
        self.response_area = QtWidgets.QPlainTextEdit()
        self.response_area.setReadOnly(True)
        self.response_area.setFont(QtGui.QFont("Courier New", 11))
        layout.addWidget(self.response_area)

        self._toggle_body("GET")

    def _toggle_body(self, method: str):
        visible = method in ("POST", "PATCH")
        self.body_label.setVisible(visible)
        self.body_input.setVisible(visible)

    def _run(self):
        method = self.method_combo.currentText()
        endpoint = self.endpoint_input.text().strip()
        if not endpoint:
            return

        body = None
        if method in ("POST", "PATCH"):
            raw = self.body_input.toPlainText().strip()
            if raw:
                try:
                    body = json.loads(raw)
                except json.JSONDecodeError as exc:
                    self.response_area.setPlainText(f"Invalid JSON body:\n{exc}")
                    return

        self.response_area.setPlainText("Running…")

        def _worker():
            try:
                result = self._api.execute(method, endpoint, body)
                text = json.dumps(result, indent=2) if result is not None else "(no content)"
                QtCore.QMetaObject.invokeMethod(
                    self, "_show_response",
                    QtCore.Qt.QueuedConnection,
                    QtCore.Q_ARG(str, text),
                )
            except (GraphError, Exception) as exc:
                QtCore.QMetaObject.invokeMethod(
                    self, "_show_response",
                    QtCore.Qt.QueuedConnection,
                    QtCore.Q_ARG(str, f"Error:\n{exc}"),
                )

        threading.Thread(target=_worker, daemon=True).start()

    @QtCore.Slot(str)
    def _show_response(self, text: str):
        self.response_area.setPlainText(text)


# ── Main application window ───────────────────────────────────────────────────

class M365ManagerApp(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Microsoft 365 Environment Manager")
        self.resize(1200, 750)

        self._stack = QtWidgets.QStackedWidget()
        self.setCentralWidget(self._stack)

        self._login_widget = LoginWidget()
        self._login_widget.logged_in.connect(self._on_logged_in)
        self._stack.addWidget(self._login_widget)

    @QtCore.Slot(object, object)
    def _on_logged_in(self, auth: M365Auth, api: GraphAPI):
        dashboard = QtWidgets.QWidget()
        d_layout = QtWidgets.QVBoxLayout(dashboard)
        d_layout.setContentsMargins(0, 0, 0, 0)
        d_layout.setSpacing(0)

        # Header bar
        header = QtWidgets.QFrame()
        header.setFixedHeight(48)
        header.setStyleSheet("background: #1e293b;")
        h_layout = QtWidgets.QHBoxLayout(header)
        h_layout.setContentsMargins(20, 0, 20, 0)

        title_lbl = QtWidgets.QLabel("Microsoft 365 Manager")
        title_lbl.setStyleSheet("font-size:15px; font-weight:600;")
        user_lbl = QtWidgets.QLabel(f"Signed in as: {auth.username}")
        user_lbl.setStyleSheet("color: #94a3b8;")

        sign_out_btn = QtWidgets.QPushButton("Sign Out")
        sign_out_btn.clicked.connect(self._sign_out)

        h_layout.addWidget(title_lbl)
        h_layout.addStretch()
        h_layout.addWidget(user_lbl)
        h_layout.addWidget(sign_out_btn)
        d_layout.addWidget(header)

        # Tabs
        tabs = QtWidgets.QTabWidget()
        tabs.addTab(UsersTab(api), "Users")
        tabs.addTab(LicensesTab(api), "Licenses")
        tabs.addTab(GraphExplorerTab(api), "Graph Explorer")
        d_layout.addWidget(tabs)

        self._stack.addWidget(dashboard)
        self._stack.setCurrentWidget(dashboard)

    def _sign_out(self):
        # Remove dashboard widget, go back to login
        if self._stack.count() > 1:
            widget = self._stack.widget(1)
            self._stack.removeWidget(widget)
            widget.deleteLater()
        self._login_widget.username_input.clear()
        self._login_widget.status_label.setText("")
        self._stack.setCurrentWidget(self._login_widget)


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    app = QtWidgets.QApplication(sys.argv)
    _apply_theme(app)
    window = M365ManagerApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
