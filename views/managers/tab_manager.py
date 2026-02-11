"""Менеджер для работы с вкладками"""
from PyQt5.QtWidgets import QMenu, QStyle
from PyQt5.QtCore import Qt
from logger import logger
from views.widgets import DetachedTabWindow
from PyQt5.QtWidgets import QApplication


class TabManager:
    """Менеджер для управления вкладками (открепление/прикрепление)"""
    
    def __init__(self, main_window):
        """
        Args:
            main_window: Ссылка на главное окно
        """
        self.main_window = main_window
    
    def show_tab_context_menu(self, position):
        """Показать контекстное меню для вкладок
        
        Args:
            position: Позиция клика относительно QTabWidget
        """
        # position - это позиция клика относительно QTabWidget
        # Проверяем, что клик был именно на tabBar, а не на содержимом вкладки
        tab_bar = self.main_window.tabs_panel.tabBar()
        tab_bar_pos = tab_bar.mapFrom(self.main_window.tabs_panel, position)
        
        # Проверяем, что позиция находится в пределах tabBar
        tab_bar_rect = tab_bar.rect()
        if not tab_bar_rect.contains(tab_bar_pos):
            # Клик был не на tabBar, а на содержимом вкладки - не показываем меню
            return
        
        tab_index = tab_bar.tabAt(tab_bar_pos)
        
        # Если не нашли вкладку по позиции, значит клик был между вкладками или вне их
        if tab_index < 0:
            return
        
        tab_name = self.main_window.tabs_panel.tabText(tab_index)
        if not tab_name:
            return
        
        menu = QMenu(self.main_window)
        
        # Проверяем, откреплена ли вкладка
        if tab_name in self.main_window.detached_windows:
            attach_action = menu.addAction("Вернуть во вкладки")
            attach_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_DialogApplyButton))
            action = menu.exec_(self.main_window.tabs_panel.mapToGlobal(position))
            if action == attach_action:
                self.attach_tab(tab_name, None)
        else:
            detach_action = menu.addAction("Открыть в отдельном окне")
            detach_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_TitleBarNormalButton))
            action = menu.exec_(self.main_window.tabs_panel.mapToGlobal(position))
            if action == detach_action:
                self.detach_tab(tab_index, tab_name)
    
    def detach_tab(self, tab_index: int, tab_name: str):
        """Открепление вкладки в отдельное окно
        
        Args:
            tab_index: Индекс вкладки
            tab_name: Название вкладки
        """
        # Получаем виджет вкладки
        tab_widget = self.main_window.tabs_panel.widget(tab_index)
        if not tab_widget:
            return
        
        # Сохраняем текущий размер виджета
        widget_size = tab_widget.size()
        
        # Удаляем вкладку из главного окна (но не удаляем сам виджет)
        self.main_window.tabs_panel.removeTab(tab_index)
        
        # Убеждаемся, что виджет видим и имеет правильный размер
        tab_widget.setParent(None)
        tab_widget.setVisible(True)
        if widget_size.isValid() and widget_size.width() > 0 and widget_size.height() > 0:
            tab_widget.resize(widget_size)
        
        # Создаем отдельное окно
        detached_window = DetachedTabWindow(tab_widget, tab_name, self.main_window)
        self.main_window.detached_windows[tab_name] = detached_window
        
        # Показываем окно
        detached_window.show()
        detached_window.raise_()
        detached_window.activateWindow()
    
    def attach_tab(self, tab_name: str, tab_widget=None):
        """Возврат вкладки в главное окно
        
        Args:
            tab_name: Название вкладки
            tab_widget: Виджет вкладки (опционально)
        """
        # Проверяем, есть ли эта вкладка в открепленных окнах
        if tab_name not in self.main_window.detached_windows:
            # Если вкладка уже не в словаре, возможно она уже была возвращена
            # Проверяем, не находится ли она уже в tabs_panel
            for i in range(self.main_window.tabs_panel.count()):
                if self.main_window.tabs_panel.tabText(i) == tab_name:
                    return
            logger.warning(f"Вкладка '{tab_name}' не найдена в detached_windows и не найдена в tabs_panel")
            return
        
        detached_window = self.main_window.detached_windows[tab_name]
        
        # Получаем виджет из окна (теперь это центральный виджет напрямую)
        if tab_widget is None:
            tab_widget = detached_window.centralWidget()
        
        if not tab_widget:
            logger.error(f"Не удалось получить виджет для вкладки '{tab_name}'")
            # Если виджет не найден, просто удаляем запись
            try:
                detached_window.setProperty("attaching", True)
                detached_window.close()
            except Exception as e:
                logger.debug(f"Не удалось закрыть открепленное окно '{tab_name}': {e}")
            if tab_name in self.main_window.detached_windows:
                del self.main_window.detached_windows[tab_name]
            return
        
        # Сохраняем размер виджета
        widget_size = tab_widget.size()
        
        # Устанавливаем флаг, чтобы closeEvent не вызывал attach_tab повторно
        detached_window.setProperty("attaching", True)
        
        # Удаляем запись из словаря перед добавлением вкладки обратно
        # Это предотвратит повторные вызовы attach_tab
        if tab_name in self.main_window.detached_windows:
            del self.main_window.detached_windows[tab_name]
        
        # Определяем позицию вкладки по имени
        tab_positions = {
            "Древовидные данные": 0,
            "Метаданные": 1,
            "Ошибки": 2,
            "Просмотр формы": 3
        }
        position = tab_positions.get(tab_name, self.main_window.tabs_panel.count())
        
        # ВАЖНО: Не удаляем виджет из окна до добавления в tabs_panel
        # QTabWidget.insertTab() автоматически установит правильного родителя
        # и удалит виджет из старого родителя
        
        # Убеждаемся, что виджет видим
        tab_widget.setVisible(True)
        
        # Восстанавливаем размер, если он был валидным
        if widget_size.isValid() and widget_size.width() > 0 and widget_size.height() > 0:
            tab_widget.resize(widget_size)
        
        # Добавляем вкладку обратно в главное окно
        # insertTab автоматически установит правильного родителя и удалит из старого
        try:
            inserted_index = self.main_window.tabs_panel.insertTab(position, tab_widget, tab_name)
            
            # Проверяем, что вкладка действительно добавлена
            if inserted_index >= 0 and inserted_index < self.main_window.tabs_panel.count():
                # Убеждаемся, что вкладка видна
                self.main_window.tabs_panel.setCurrentIndex(inserted_index)
                self.main_window.tabs_panel.setTabVisible(inserted_index, True)
                
                # Теперь можно удалить виджет из окна, так как он уже в tabs_panel
                try:
                    detached_window.setCentralWidget(None)
                except Exception as e:
                    logger.warning(f"Ошибка при удалении центрального виджета: {e}")
                
                # Принудительно обновляем отображение
                tab_widget.show()
                tab_widget.update()
                self.main_window.tabs_panel.update()
                
                # Принудительно перерисовываем
                QApplication.processEvents()
            else:
                logger.error(f"Ошибка: вкладка не была добавлена правильно. inserted_index={inserted_index}, count={self.main_window.tabs_panel.count()}")
        except Exception as e:
            logger.error(f"Ошибка при добавлении вкладки в tabs_panel: {e}", exc_info=True)
        
        # Закрываем окно после того, как вкладка добавлена
        try:
            detached_window.close()
        except Exception as e:
            logger.warning(f"Ошибка при закрытии окна: {e}")
        
        logger.info(f"Вкладка '{tab_name}' успешно возвращена в главное окно на позицию {position}")
