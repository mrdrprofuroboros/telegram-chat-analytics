import pandas as pd
from datetime import datetime, timedelta
import json
import openpyxl

# В начале файла добавим константу с форматом даты
DATE_FORMAT = '%Y-%m-%dT%H:%M:%S'
MY_ID = 'user44366287'

def get_message_text(message):
    """Извлекает текст сообщения, учитывая разные форматы"""
    if isinstance(message['text'], str):
        return message['text'].lower()
    elif isinstance(message['text'], list):
        # Если текст является списком, объединяем все текстовые элементы
        return ' '.join(str(item) if isinstance(item, str) else item.get('text', '') 
                       for item in message['text']).lower()
    return ''

def analyze_friendship_metrics(data, output_file='friends_metrics.xlsx'):
    results = []
    
    for chat in data['chats']['list']:
        if chat['type'] != 'personal_chat':
            continue
            
        chat_name = chat['name']
        chat_id = chat['id']
        messages = [m for m in chat['messages'] if m['type'] == 'message']
        
        if not messages:
            continue
            
        # Создаем DataFrame для удобной работы с датами
        msgs_df = pd.DataFrame(messages)
        msgs_df['date'] = pd.to_datetime(msgs_df['date'], format=DATE_FORMAT)
        msgs_df['year'] = msgs_df['date'].dt.year
        years = sorted(msgs_df['year'].unique())
        
        # Для каждого года и для всего периода
        periods = years + ['all']
        
        for period in periods:
            # Фильтруем сообщения по периоду
            if period == 'all':
                period_messages = messages
                period_df = msgs_df
            else:
                period_df = msgs_df[msgs_df['year'] == period]
                period_messages = [m for m in messages 
                                 if datetime.fromisoformat(m['date']).year == period]
            
            if not period_messages:
                continue

            # Получаем полное имя собеседника из его сообщений
            their_messages = [m for m in period_messages if m['from_id'] != MY_ID]
            if their_messages:
                chat_name = their_messages[0]['from']  # Берем имя из первого сообщения собеседника
            
            # Базовые метрики
            total_messages = len(period_messages)
            my_messages = len([m for m in period_messages if m['from_id'] == MY_ID])
            their_messages = total_messages - my_messages
            
            # Временные метрики
            dates = pd.to_datetime([m['date'] for m in period_messages], format=DATE_FORMAT)
            chat_duration = (dates.max() - dates.min()).days
            messages_per_day = total_messages / max(chat_duration, 1)
            
            # Анализ времени ответа
            response_times = []
            for i in range(1, len(period_messages)):
                if period_messages[i]['from_id'] != period_messages[i-1]['from_id']:
                    time1 = datetime.fromisoformat(period_messages[i-1]['date'])
                    time2 = datetime.fromisoformat(period_messages[i]['date'])
                    diff_minutes = (time2 - time1).total_seconds() / 60
                    if diff_minutes < 60 * 24:  # Учитываем только ответы в течение суток
                        response_times.append(diff_minutes)
            
            avg_response_time = sum(response_times) / len(response_times) if response_times else 0
            
            # Анализ длины сообщений
            my_messages_texts = [m['text'] for m in period_messages if m['from_id'] == MY_ID]
            their_messages_texts = [m['text'] for m in period_messages if m['from_id'] != MY_ID]
            
            my_avg_length = sum(len(text) for text in my_messages_texts) / len(my_messages_texts) if my_messages_texts else 0
            their_avg_length = sum(len(text) for text in their_messages_texts) / len(their_messages_texts) if their_messages_texts else 0
            
            # Эмоциональный анализ
            emoji_count = sum(1 for m in period_messages if '😊' in m['text'] or '😂' in m['text'] or '❤️' in m['text'])
            
            # Анализ времени суток
            night_messages = sum(1 for m in period_messages if 0 <= datetime.fromisoformat(m['date']).hour < 6)
            
            # Анализ инициаций
            period_messages.sort(key=lambda x: x['date_unixtime'])
            my_initiations = 0
            other_initiations = 0
            last_message_time = None
            
            for msg in period_messages:
                current_time = datetime.fromtimestamp(int(msg['date_unixtime']))
                if last_message_time is None or (current_time - last_message_time > timedelta(hours=6)):
                    if msg['from_id'] == MY_ID:
                        my_initiations += 1
                    else:
                        other_initiations += 1
                last_message_time = current_time
            
            # Первые сообщения
            first_messages = []
            for i, msg in enumerate(period_messages[:6]):
                sender = "Я: " if msg['from_id'] == MY_ID else f"{msg['from']}: "
                first_messages.append(f"{sender}{msg['text']}")
            chat_start = " | ".join(first_messages)
            
            # Добавляем анализ "спасибо"
            thanks_words = ['спасибо', 'благодарю', 'thanks', 'thank you', 'thx']
            my_thanks = sum(1 for m in period_messages 
                          if m['from_id'] == MY_ID and 
                          any(word in get_message_text(m) for word in thanks_words))
            their_thanks = sum(1 for m in period_messages 
                             if m['from_id'] != MY_ID and 
                             any(word in get_message_text(m) for word in thanks_words))
            
            # Обновляем словарь results с новыми метриками
            results.append({
                'chat_name': chat_name,
                'period': period,
                'total_messages': total_messages,
                'initiation_ratio': round(my_initiations / max(other_initiations, 1), 2),
                'avg_response_time_min': round(avg_response_time, 2),
                'thanks_ratio': round(my_thanks / max(their_thanks, 1), 2),
                'messages_per_day': round(messages_per_day, 2),
                'message_ratio': round(my_messages / max(their_messages, 1), 2),
                'my_initiations': my_initiations,
                'other_initiations': other_initiations,
                'my_avg_msg_length': round(my_avg_length, 2),
                'their_avg_msg_length': round(their_avg_length, 2),
                'night_messages_percent': round(night_messages * 100 / total_messages, 2),
                'my_messages_count': my_messages,
                'their_messages_count': their_messages,
                'chat_id': chat_id,
                'my_thanks_count': my_thanks,
                'their_thanks_count': their_thanks
            })
    
    df = pd.DataFrame(results)
    if len(df) > 0:
        df = df.sort_values(['total_messages', 'period'], ascending=False)
        
        # Сохраняем в Excel с автонастройкой ширины колонок
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            worksheet = writer.sheets['Sheet1']
            for idx, col in enumerate(df.columns):
                # Устанавливаем ширину колонки равной длине заголовка
                worksheet.column_dimensions[openpyxl.utils.get_column_letter(idx + 1)].width = len(str(col)) + 2
        
        print(f'Анализ сохранен в файл: {output_file}')
        return df
    return df

if __name__ == "__main__":
    # Загрузка данных из JSON файла
    try:
        with open('result.json', 'r', encoding='utf-8') as file:
            data = json.load(file)
        
        # Анализ и сохранение результатов в Excel
        analyze_friendship_metrics(data, 'friends_metrics_2024.xlsx')
        
    except FileNotFoundError:
        print("Ошибка: Файл 'result.json' не найден!")
    except json.JSONDecodeError:
        print("Ошибка: Невозможно прочитать JSON файл!")
    except Exception as e:
        print(f"Произошла ошибка: {str(e)}") 