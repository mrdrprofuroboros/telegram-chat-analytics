import os
import pandas as pd
from datetime import datetime, timedelta
import json
import openpyxl
import traceback
from tqdm import tqdm
import matplotlib.pyplot as plt
import matplotlib.animation as animation
import numpy as np

import plotly.graph_objects as go
from plotly.subplots import make_subplots

# В начале файла добавим константу с форматом даты
DATE_FORMAT = '%Y-%m-%dT%H:%M:%S'
MY_ID = 'user44366287'

def get_message_text(message):
    """Извлекает текст сообщения, учитывая разные форматы"""
    if isinstance(message, str):
        return message.lower()
    elif isinstance(message, list):
        # Если текст является списком, объединяем все текстовые элементы
        return ' '.join(str(item) if isinstance(item, str) else item.get('text', '') 
                       for item in message).lower()
    elif isinstance(message, dict):
        return str(message.get('text', '')).lower()
    return ''


def create_interactive_chart(df):
    """Creates an interactive HTML chart with a slider and stacked bars for group members"""
    df['period'] = pd.to_datetime(df['period'].astype(str))
    unique_periods = sorted(df['period'].unique())
    
    # Create color palette for members (excluding 'My Messages' color)
    nbars = 12
    crop = 24
    # Use qualitative colormap from matplotlib
    cmap = plt.cm.Set3
    
    # Create the figure
    fig = go.Figure()
    
    # Process each period
    current_trace_idx = 0  # Keep track of current trace index
    period_start_indices = {}  # Store start index for each period

    for period_idx, period in tqdm(enumerate(unique_periods), desc="Processing periods"):
        period_df = df[df['period'] == period]
        top_chats = (period_df.groupby('base_chat')['my_messages_count']
                    .first()
                    .nlargest(nbars)
                    .index[::-1])

        period_start_indices[period] = current_trace_idx
        
        # Add empty traces to reach nbars if needed
        empty_slots_needed = nbars - len(top_chats)
        for i in range(empty_slots_needed):
            fig.add_trace(
                go.Bar(
                    x=[0],
                    y=[f'{nbars-i}'],
                    orientation='h',
                    visible=period_idx == 0,
                    showlegend=False,
                    hoverinfo='skip',
                    marker_color='rgba(0,0,0,0)'
                )
            )
            current_trace_idx += 1
    
        # Process actual chats
        for chat in top_chats:
            chat_df = period_df[period_df['base_chat'] == chat]
            
            # My messages bar
            fig.add_trace(
                go.Bar(
                    x=[chat_df[chat_df['member_name'] == 'Me']['my_messages_count'].iloc[0]],
                    y=[chat[:crop]],
                    orientation='h',
                    visible=period_idx == 0,
                    name='My Messages',
                    marker_color='#FF4444',
                    hovertemplate='My Messages: %{x}<extra></extra>',
                    showlegend=False,
                )
            )
            current_trace_idx += 1
            
            # Other members bars
            other_members = (chat_df[chat_df['member_name'] != 'Me']
                           .sort_values('their_messages_count', ascending=False))
            
            # Add member traces
            for idx, (_, member) in enumerate(other_members.iterrows()):
                color_rgba = cmap(idx % 16)
                color_hex = f'rgba({int(color_rgba[0]*255)},{int(color_rgba[1]*255)},{int(color_rgba[2]*255)},{color_rgba[3]})'
                
                fig.add_trace(
                    go.Bar(
                        x=[member['their_messages_count']],
                        y=[chat[:crop]],
                        orientation='h',
                        visible=period_idx == 0,
                        name=member['member_name'],
                        marker_color=color_hex,
                        hovertemplate=f'{member["member_name"]}: %{{x}}<extra></extra>',
                        showlegend=False,
                    )
                )
                current_trace_idx += 1
        
    # Create slider steps
    steps = []
    for i, period in enumerate(unique_periods):
        visible = [False] * len(fig.data)
        start_idx = period_start_indices[period]
        
        # Get end index from next period's start index, or use total length for last period
        if i < len(unique_periods) - 1:
            end_idx = period_start_indices[unique_periods[i + 1]]
        else:
            end_idx = len(fig.data)
            
        visible[start_idx:end_idx] = [True] * (end_idx - start_idx)
        
        steps.append(dict(
            method="update",
            args=[
                {"visible": visible},
                {"title": f"Messages by Chat ({period.strftime('%Y-%m')})"}
            ],
            label=period.strftime("%Y-%m")
        ))
    
    sliders = [dict(
        active=0,
        currentvalue={"prefix": "Date: "},
        pad={"t": 50},
        steps=steps
    )]
    
    # Update layout
    fig.update_layout(
        sliders=sliders,
        title="Messages by Chat",
        xaxis_title="Messages per Month",
        xaxis_range=[0, 3300],
        height=650,
        showlegend=True,
        barmode='stack',
        yaxis={
            'autorange': True,
            'showticklabels': True,
            'tickfont': {'size': 12}
        },
        uniformtext=dict(minsize=8, mode='hide'),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
        margin=dict(l=200),
        plot_bgcolor='white',
        paper_bgcolor='white',
        hoverlabel=dict(bgcolor='white'),
    )
    
    fig.update_xaxes(
        showgrid=True,
        gridwidth=1,
        gridcolor='LightGrey'
    )
    
    fig.update_yaxes(
        showgrid=True,
        gridwidth=1,
        gridcolor='LightGrey'
    )
    
    fig.write_html('chat_evolution.html')


def analyze_friendship_metrics(data, output_file='friends_metrics.xlsx', cache_file='chat_metrics.pkl'):
    if os.path.exists(cache_file):
        df = pd.read_pickle(cache_file)
        create_interactive_chart(df)
        return df

    results = []
    for chat in tqdm(data['chats']['list']):
        if chat['type'] == 'saved_messages' or not chat.get('messages'):
            continue
            
        # Create DataFrame for all messages
        msgs_df = pd.DataFrame([m for m in chat['messages'] if m['type'] == 'message'])
        if len(msgs_df) <= 1:
            continue
            
        # Basic date processing
        msgs_df['date'] = pd.to_datetime(msgs_df['date'])
        msgs_df['period'] = msgs_df['date'].dt.to_period('M')
        msgs_df['hour'] = msgs_df['date'].dt.hour
        msgs_df['is_night'] = msgs_df['hour'].between(0, 5)
        msgs_df['is_my_message'] = msgs_df['from_id'] == MY_ID
        
        # Handle text processing
        msgs_df['message_text'] = msgs_df['text'].apply(get_message_text)
        msgs_df['msg_length'] = msgs_df['message_text'].str.len()
        
        # Calculate time differences for initiations
        msgs_df['time_diff'] = msgs_df['date'].diff()
        msgs_df['is_new_conversation'] = msgs_df['time_diff'] > pd.Timedelta(hours=6)
        
        # Group by period
        for period, period_df in msgs_df.groupby('period'):
            # Base metrics
            total_messages = len(period_df)
            chat_duration = (period_df['date'].max() - period_df['date'].min()).days or 1
            
            # My messages
            my_messages = period_df[period_df['is_my_message']]
            my_messages_count = len(my_messages)
            
            # Initiations
            my_initiations = len(period_df[period_df['is_new_conversation'] & period_df['is_my_message']])
            other_initiations = len(period_df[period_df['is_new_conversation'] & ~period_df['is_my_message']])
            
            # Night messages
            night_messages = len(period_df[period_df['is_night']])
            
            # Base result dictionary
            base_result = {
                'period': period,
                'base_chat': chat['name'],
                'chat_type': chat['type'],
                'chat_id': chat['id'],
                'total_messages': total_messages,
                'my_messages_count': my_messages_count,
                'messages_per_day': round(total_messages / chat_duration, 2),
                'night_messages_count': night_messages,
                'my_initiations': my_initiations,
                'other_initiations': other_initiations,
            }
            
            # Add my messages entry
            my_result = base_result.copy()
            my_result.update({
                'member_name': 'Me',
                'their_messages_count': 0,
                'my_avg_msg_length': round(my_messages['msg_length'].mean(), 2) if my_messages_count else 0,
            })
            results.append(my_result)
            
            # Process other members
            other_members = period_df[~period_df['is_my_message']].groupby('from')
            for member_name, member_msgs in other_members:
                member_result = base_result.copy()
                member_result.update({
                    'member_name': member_name,
                    'their_messages_count': len(member_msgs),
                    'their_avg_msg_length': round(member_msgs['msg_length'].mean(), 2),
                })
                results.append(member_result)
    
    df = pd.DataFrame(results)
    
    # Sort and save
    if len(df) > 0:
        df = df.sort_values(['my_messages_count', 'period'], ascending=False)
        df.to_pickle(cache_file)
        create_interactive_chart(df)
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            worksheet = writer.sheets['Sheet1']
            for idx, col in enumerate(df.columns):
                worksheet.column_dimensions[openpyxl.utils.get_column_letter(idx + 1)].width = len(str(col)) + 2
    
    return df

if __name__ == "__main__":
    # Загрузка данных из JSON файла
    try:
        cache_file = 'chat_metrics.pkl'
        if os.path.exists(cache_file):
            df = pd.read_pickle(cache_file)
            create_interactive_chart(df)
        else:
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
        traceback.print_exc()