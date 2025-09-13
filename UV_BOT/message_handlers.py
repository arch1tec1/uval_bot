from telegram import Update
from telegram.ext import CallbackContext
from config import (
    TARGET_TOPIC_ID, FORM_KEYWORDS, COMMAND_KEYWORDS, 
    OFFICER_DICT, NAMESAKES, SPECIAL_USER_ID
)
from user_manager import UserManager
from document_handler import DocumentHandler

class MessageHandlers:

    def __init__(self):
        self.user_manager = UserManager()
        self.doc_handler = DocumentHandler()
        self.fio_data_list, self.fio_by_last_name = self.doc_handler.create_fio_data_structure(
            self.doc_handler.get_keywords_fio()
        )
    
    async def handle_message(self, update: Update, context: CallbackContext):

        message_text = update.message.text.lower()
        message_words = set(message_text.split())
        message_words_list = message_text.split()
        chat_id = update.message.chat_id
        human_id = str(update.message.from_user.id)
        topic_id = update.message.message_thread_id
        
        print(f"Topic ID: {topic_id}")
        print(f"Сообщение: {update.message.text}")
        
        # Проверка правильного чата
        if topic_id != TARGET_TOPIC_ID:
            return
        
        # Проверка смены дня и обновление документа
        await self._check_and_update_daily_document()
        
        # Обработка различных команд
        if FORM_KEYWORDS & message_words:
            await self._handle_add_to_list(update, message_words, message_words_list)
        elif "убрать" in message_words:
            await self._handle_remove_from_list(update, message_words_list)
        elif "документ" in message_words:
            await self._handle_document_request(update, context, message_words, human_id)
        elif "список" in message_words:
            await self._handle_list_request(update, context, message_words, human_id, chat_id, topic_id)
        elif "лист" in message_words:
            await self._handle_text_list_request(update, context, chat_id, topic_id)
        elif "офицер" in message_words:
            await self._handle_officer_change(update, message_words, human_id)
        elif "очистить" in message_words:
            await self._handle_clear_list(update, human_id)
        elif not (COMMAND_KEYWORDS & message_words) and human_id != SPECIAL_USER_ID:
            await update.message.reply_text(
                'А ничо тот факт, что я не нашел в твоем сообщении ключевых слов. '
                'Скорее всего ты написал не в тот чат. Поменяй чат, либо напиши корректно.'
            )
    
    async def _check_and_update_daily_document(self):
        """Проверка и обновление документа при смене дня"""
        # Проверка даты в документе
        current_date = self.doc_handler.get_today_date_info()[0].split()[-4:]
        doc_date = self.doc_handler.get_document_date()
        
        if doc_date != current_date:
            # Очищаем списки при смене дня
            self.user_manager.clear_lists()
            self.doc_handler.clean_table()
            self.doc_handler.update_date_header()
    
    async def _handle_add_to_list(self, update: Update, message_words, message_words_list):
        """Обработка добавления в список"""
        added = False
        povtor = False
        
        for word in message_words_list:
            if word in self.fio_by_last_name:
                last_name = word
                for fio_entry in self.fio_by_last_name[last_name]:
                    fio = fio_entry['full_name']
                    initials = fio_entry['initials']
                    
                    # Проверка однофамильцев
                    if last_name in NAMESAKES:
                        index = message_words_list.index(last_name)
                        if index + 1 < len(message_words_list):
                            next_word_initial = message_words_list[index + 1][0]
                            message_initials = f"{last_name} {next_word_initial}"
                            if message_initials == initials:
                                form_type = self.user_manager.get_form_type(message_words)
                                if self.user_manager.add_to_list(fio, form_type, self.fio_by_last_name):
                                    await update.message.reply_text("Ты добавлен!")
                                    added = True
                                    break
                                else:
                                    await update.message.reply_text(
                                        "А ничо тот факт, что ты добавляешься не первый раз за сегодня?"
                                    )
                                    povtor = True
                                    break
                    else:
                        form_type = self.user_manager.get_form_type(message_words)
                        if self.user_manager.add_to_list(fio, form_type, self.fio_by_last_name):
                            await update.message.reply_text("Ты добавлен!")
                            added = True
                            break
                        else:
                            await update.message.reply_text(
                                "А ничо тот факт, что ты добавляешься не первый раз за сегодня?"
                            )
                            povtor = True
                            break
        
        if not added and not povtor:
            await update.message.reply_text(
                "А ничо тот факт, что данные введены некорректно."
            )
    
    async def _handle_remove_from_list(self, update: Update, message_words_list):
        """Обработка удаления из списка"""
        removed = False
        
        for word in message_words_list:
            if word in self.fio_by_last_name:
                last_name = word
                for fio_entry in self.fio_by_last_name[last_name]:
                    fio = fio_entry['full_name']
                    first_initial = fio_entry['initials'][-1]
                    
                    # Проверка однофамильцев
                    if last_name in NAMESAKES:
                        index_initial = message_words_list.index(last_name) + 1
                        first_initial_in_message = message_words_list[index_initial]
                        if first_initial_in_message == first_initial:
                            if not self.user_manager.remove_from_list(fio):
                                await update.message.reply_text(
                                    "А ничо тот факт, что тебя и так нет в списке?"
                                )
                            else:
                                await update.message.reply_text("Убрал тебя из списка.")
                                removed = True
                                break
                    else:
                        if not self.user_manager.remove_from_list(fio):
                            await update.message.reply_text(
                                "А ничо тот факт, что тебя и так нет в списке?"
                            )
                        else:
                            await update.message.reply_text("Убрал тебя из списка.")
                            removed = True
                            break
        
        if not removed:
            await update.message.reply_text(
                "А ничо тот факт, что ты ввел данные некорректно?"
            )
    
    async def _handle_document_request(self, update: Update, context: CallbackContext, message_words, human_id):
        """Обработка запроса документа"""
        if not self.user_manager.is_main_admin(human_id):
            await update.message.reply_text("Недостаточно прав!")
            return
        
        # Обработка различных типов документов
        if {"спортмасс", "спортмасса"} & message_words:
            self.doc_handler.insert_data_into_table(
                self.user_manager.spisok_v_uval_fio, 
                self.user_manager.spisok_v_uval_forma, 
                'на спортивно-массовую работу', 
                self.user_manager.dismissal_time
            )
        elif "город" in message_words:
            self.doc_handler.insert_data_into_table(
                self.user_manager.spisok_v_uval_fio, 
                self.user_manager.spisok_v_uval_forma, 
                'в увольнение', 
                self.user_manager.dismissal_time
            )
        elif "культпоход" in message_words:
            self.doc_handler.insert_data_into_table(
                self.user_manager.spisok_v_uval_fio, 
                self.user_manager.spisok_v_uval_forma, 
                'в культпоход', 
                self.user_manager.dismissal_time
            )
        elif "время" in message_words:
            time_words = sorted(list(message_words))
            self.user_manager.update_dismissal_time(time_words[0])
            self.doc_handler.insert_data_into_table(
                self.user_manager.spisok_v_uval_fio, 
                self.user_manager.spisok_v_uval_forma, 
                vremya=self.user_manager.dismissal_time
            )
        else:
            self.doc_handler.insert_data_into_table(
                self.user_manager.spisok_v_uval_fio, 
                self.user_manager.spisok_v_uval_forma
            )
        
        await context.bot.send_document(chat_id=human_id, document=open("Новый список 1.docx", 'rb'))
        await update.message.reply_text("Посмотри личные сообщения.")
    
    async def _handle_list_request(self, update: Update, context: CallbackContext, message_words, human_id, chat_id, topic_id):
        """Обработка запроса списка"""
        if not self.user_manager.can_send_list(human_id):
            await context.bot.send_message(chat_id=chat_id, message_thread_id=topic_id, 
                                         text=self.user_manager.format_list_for_display(self.doc_handler))
            return
        
        # Обработка различных типов списков (аналогично документу)
        if {"спортмасс", "спортмасса"} & message_words:
            self.doc_handler.insert_data_into_table(
                self.user_manager.spisok_v_uval_fio, 
                self.user_manager.spisok_v_uval_forma, 
                'на спортивно-массовую работу', 
                self.user_manager.dismissal_time
            )
        elif "город" in message_words:
            self.doc_handler.insert_data_into_table(
                self.user_manager.spisok_v_uval_fio, 
                self.user_manager.spisok_v_uval_forma, 
                'в увольнение', 
                self.user_manager.dismissal_time
            )
        elif "культпоход" in message_words:
            self.doc_handler.insert_data_into_table(
                self.user_manager.spisok_v_uval_fio, 
                self.user_manager.spisok_v_uval_forma, 
                'в культпоход', 
                self.user_manager.dismissal_time
            )
        elif "время" in message_words:
            time_words = sorted(list(message_words))
            self.user_manager.update_dismissal_time(time_words[0])
            self.doc_handler.insert_data_into_table(
                self.user_manager.spisok_v_uval_fio, 
                self.user_manager.spisok_v_uval_forma, 
                vremya=self.user_manager.dismissal_time
            )
        else:
            self.doc_handler.insert_data_into_table(
                self.user_manager.spisok_v_uval_fio, 
                self.user_manager.spisok_v_uval_forma
            )
        
        await context.bot.send_document(chat_id=chat_id, message_thread_id=topic_id, 
                                      document=open("Новый список 1.docx", 'rb'))
    
    async def _handle_text_list_request(self, update: Update, context: CallbackContext, chat_id, topic_id):
        """Обработка запроса текстового списка"""
        await context.bot.send_message(chat_id=chat_id, message_thread_id=topic_id, 
                                     text=self.user_manager.format_list_for_display(self.doc_handler))
    
    async def _handle_officer_change(self, update: Update, message_words, human_id):
        """Обработка смены офицера"""
        if not self.user_manager.is_admin(human_id):
            await update.message.reply_text(
                "А ничо тот факт, что ты не можешь заменить офицера в списке?"
            )
            return
        
        for key, value in OFFICER_DICT.items():
            key_parts = key.split('-')
            for z in key_parts:
                if z in message_words:
                    doljnost, zvanie, oficer_fio = value
                    self.doc_handler.change_officer(doljnost, zvanie, oficer_fio)
                    await update.message.reply_text("Заменил офицера.")
                    break
    
    async def _handle_clear_list(self, update: Update, human_id):
        """Обработка очистки списка"""
        if not self.user_manager.is_admin(human_id):
            await update.message.reply_text(
                "А ничо тот факт, что ты не можешь очистить список?"
            )
            return
        
        self.user_manager.clear_lists()
        self.doc_handler.clean_table()
        await update.message.reply_text("Список полностью очищен.")