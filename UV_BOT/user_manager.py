from datetime import datetime
import pytz
from config import ADMIN_LIST, MAIN_ADMIN_LIST, SPECIAL_USER_ID, NAMESAKES, DEFAULT_DISMISSAL_TIME

class UserManager:
    
    def __init__(self):
        self.spisok_v_uval_forma = []
        self.spisok_v_uval_fio = []
        self.dismissal_time = DEFAULT_DISMISSAL_TIME
        self.moscow_tz = pytz.timezone('Europe/Moscow')
    
    def is_admin(self, user_id):
        """Проверка, является ли пользователь администратором"""
        return str(user_id) in ADMIN_LIST.values()
    
    def is_main_admin(self, user_id):
        """Проверка, является ли пользователь главным администратором"""
        return str(user_id) in MAIN_ADMIN_LIST.values()
    
    def is_special_user(self, user_id):
        """Проверка, является ли пользователь специальным пользователем"""
        return str(user_id) == SPECIAL_USER_ID
    
    def can_send_list(self, user_id):
        """Проверка, может ли пользователь запросить список"""
        current_time = datetime.now(self.moscow_tz)
        return (self.is_admin(user_id) and (current_time.hour > 18 or (current_time.hour == 18 and current_time.minute >= 30))) or self.is_special_user(user_id)
    
    def add_to_list(self, fio, form_type, fio_by_last_name):
        """Добавление пользователя в список"""
        if fio not in self.spisok_v_uval_fio:
            self.spisok_v_uval_fio.append(fio)
            self.spisok_v_uval_forma.append(f"{fio} - {form_type}")
            return True
        return False
    
    def remove_from_list(self, fio):
        """Удаление пользователя из списка"""
        if fio in self.spisok_v_uval_fio:
            index = self.spisok_v_uval_fio.index(fio)
            self.spisok_v_uval_fio.pop(index)
            self.spisok_v_uval_forma.pop(index)
            return True
        return False
    
    def clear_lists(self):
        """Очистка всех списков"""
        self.spisok_v_uval_forma.clear()
        self.spisok_v_uval_fio.clear()
        self.dismissal_time = DEFAULT_DISMISSAL_TIME
    
    def get_sorted_lists(self):
        """Получение отсортированных списков"""
        sorted_fio = sorted(self.spisok_v_uval_fio)
        sorted_forma = sorted(self.spisok_v_uval_forma)
        return sorted_fio, sorted_forma
    
    def format_list_for_display(self, doc_handler):
        """Форматирование списка для отображения"""
        sorted_fio, sorted_forma = self.get_sorted_lists()
        spisok_lines = [f"{i + 1}. {fio_forma[:fio_forma.find('-')].title()}{fio_forma[fio_forma.find('-'):]}" 
                       for i, fio_forma in enumerate(sorted_forma)]
        spisok = '\n'.join(spisok_lines)
        
        date_info = doc_handler.get_today_date_info(vremya=self.dismissal_time)
        responsible = doc_handler.get_responsible_officer()
        
        return (f"Список {date_info[0]}:"
                f"\n\n{spisok}\n\n"
                f"Время увольнения: {date_info[1]}"
                f"\n\nОтветственный сегодня - {responsible}")
    
    def find_user_by_message(self, message_words_list, fio_by_last_name):
        """Поиск пользователя по сообщению"""
        for word in message_words_list:
            if word in fio_by_last_name:
                last_name = word
                for fio_entry in fio_by_last_name[last_name]:
                    fio = fio_entry['full_name']
                    initials = fio_entry['initials']
                    
                    # Проверка однофамильцев
                    if last_name in NAMESAKES:
                        index = message_words_list.index(last_name)
                        if index + 1 < len(message_words_list):
                            next_word_initial = message_words_list[index + 1][0]
                            message_initials = f"{last_name} {next_word_initial}"
                            if message_initials == initials:
                                return fio
                    else:
                        return fio
        return None
    
    def get_form_type(self, message_words):
        """Определение типа формы одежды"""
        if 'спорт' in message_words:
            return 'спортивная ФО'
        elif 'офиска' in message_words:
            return 'офисная ФО'
        elif 'гражданка' in message_words:
            return 'гражданская ФО'
        return None
    
    def update_dismissal_time(self, time_string):
        """Обновление времени увольнения"""
        self.dismissal_time = time_string