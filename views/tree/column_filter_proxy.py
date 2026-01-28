"""
Прокси-модель, показывающая только видимые столбцы.
Устраняет проблему «пустого места» при скрытии столбцов: вид получает columnCount = число видимых,
поэтому скрытые столбцы не занимают место.
"""
from PyQt5.QtCore import QAbstractProxyModel, QModelIndex, Qt


class ColumnFilterProxyModel(QAbstractProxyModel):
    """
    Прокси над источником: отдаёт только видимые столбцы (логические индексы 0, 1, 2, ...).
    В createIndex передаётся internalPointer источника, чтобы QTreeWidget.itemFromIndex() работал.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._visible_columns = []  # список логических индексов столбцов источника
        # Ключ (internalPointer, row, proxy_col) -> source QModelIndex для mapToSource
        self._index_cache = {}

    def setSourceModel(self, model):
        old = self.sourceModel()
        if old:
            old.columnsInserted.disconnect(self._on_source_columns_changed)
            old.columnsRemoved.disconnect(self._on_source_columns_changed)
            old.modelReset.disconnect(self._invalidate_cache)
        super().setSourceModel(model)
        if model:
            model.columnsInserted.connect(self._on_source_columns_changed)
            model.columnsRemoved.connect(self._on_source_columns_changed)
            model.modelReset.connect(self._invalidate_cache)
            if not self._visible_columns:
                self._visible_columns = list(range(model.columnCount()))
        self._invalidate_cache()

    def _on_source_columns_changed(self, parent, first, last):
        self._invalidate_cache()
        # Подстроить _visible_columns под новый columnCount источника
        src_count = self.sourceModel().columnCount(parent) if parent.isValid() else self.sourceModel().columnCount()
        self._visible_columns = [c for c in self._visible_columns if c < src_count]
        for c in range(first, min(last + 1, src_count)):
            if c not in self._visible_columns:
                self._visible_columns.append(c)
        self._visible_columns.sort()
        self.layoutChanged.emit()

    def _invalidate_cache(self):
        self._index_cache.clear()

    def visibleColumns(self):
        """Список логических индексов видимых столбцов в источнике."""
        return list(self._visible_columns)

    def setVisibleColumns(self, source_indices):
        """Установить видимые столбцы по логическим индексам источника. Эмитит layoutChanged."""
        src = self.sourceModel()
        if not src:
            return
        count = src.columnCount()
        self._visible_columns = [c for c in source_indices if 0 <= c < count]
        self._invalidate_cache()
        self.layoutChanged.emit()

    def isSourceColumnVisible(self, source_column):
        return source_column in self._visible_columns

    def setSourceColumnVisible(self, source_column, visible):
        if visible and source_column not in self._visible_columns:
            self._visible_columns.append(source_column)
            self._visible_columns.sort()
            self._invalidate_cache()
            self.layoutChanged.emit()
        elif not visible and source_column in self._visible_columns:
            self._visible_columns.remove(source_column)
            self._invalidate_cache()
            self.layoutChanged.emit()

    def columnCount(self, parent=QModelIndex()):
        if not self.sourceModel():
            return 0
        return len(self._visible_columns)

    def rowCount(self, parent=QModelIndex()):
        return self.sourceModel().rowCount(self.mapToSource(parent)) if self.sourceModel() else 0

    def _cache_key(self, proxy_index):
        # internalPointer — указатель на элемент (QTreeWidgetItem); используем id() для ключа кэша
        ptr = proxy_index.internalPointer()
        return (id(ptr) if ptr is not None else 0, proxy_index.row(), proxy_index.column())

    def index(self, row, column, parent=QModelIndex()):
        if not self.sourceModel() or row < 0 or column < 0 or column >= len(self._visible_columns):
            return QModelIndex()
        source_parent = self.mapToSource(parent)
        source_col = self._visible_columns[column]
        source_index = self.sourceModel().index(row, source_col, source_parent)
        if not source_index.isValid():
            return QModelIndex()
        ptr = source_index.internalPointer()
        key = (id(ptr) if ptr is not None else 0, row, column)
        self._index_cache[key] = source_index
        return self.createIndex(row, column, ptr)

    def parent(self, index):
        if not index.isValid():
            return QModelIndex()
        key = self._cache_key(index)
        source_index = self._index_cache.get(key)
        if source_index is None:
            return QModelIndex()
        source_parent = source_index.parent()
        return self.mapFromSource(source_parent)

    def mapToSource(self, proxy_index):
        if not proxy_index.isValid():
            return QModelIndex()
        key = self._cache_key(proxy_index)
        return self._index_cache.get(key, QModelIndex())

    def mapFromSource(self, source_index):
        if not source_index.isValid():
            return QModelIndex()
        if source_index.column() not in self._visible_columns:
            return QModelIndex()
        proxy_col = self._visible_columns.index(source_index.column())
        ptr = source_index.internalPointer()
        key = (id(ptr) if ptr is not None else 0, source_index.row(), proxy_col)
        self._index_cache[key] = source_index
        return self.createIndex(source_index.row(), proxy_col, ptr)

    def data(self, index, role=Qt.DisplayRole):
        src = self.mapToSource(index)
        return self.sourceModel().data(src, role) if src.isValid() else None

    def setData(self, index, value, role=Qt.EditRole):
        src = self.mapToSource(index)
        return self.sourceModel().setData(src, value, role) if src.isValid() else False

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if not self.sourceModel() or orientation != Qt.Horizontal or section < 0 or section >= len(self._visible_columns):
            return None
        return self.sourceModel().headerData(self._visible_columns[section], orientation, role)

    def flags(self, index):
        src = self.mapToSource(index)
        return self.sourceModel().flags(src) if src.isValid() else Qt.NoItemFlags

    def insertRows(self, row, count, parent=QModelIndex()):
        return self.sourceModel().insertRows(row, count, self.mapToSource(parent))

    def removeRows(self, row, count, parent=QModelIndex()):
        return self.sourceModel().removeRows(row, count, self.mapToSource(parent))

    def hasChildren(self, parent=QModelIndex()):
        return self.sourceModel().hasChildren(self.mapToSource(parent)) if self.sourceModel() else False
