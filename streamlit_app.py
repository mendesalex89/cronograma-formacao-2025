import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import locale

# Set page config for wide layout
st.set_page_config(page_title="Cronograma de Formação 2025", layout="wide")

# Set locale to Portuguese
try:
    locale.setlocale(locale.LC_TIME, 'pt_PT.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
    except locale.Error:
        pass

def get_start_date_from_week(week, year=2025):
    """Returns the Monday of the given week number."""
    # Week 1 is the first week with a Thursday.
    # Note: '1' is Monday in %w
    d = datetime.strptime(f'{year}-W{int(week)}-1', "%Y-W%W-%w")
    return d

def generate_schedule():
    st.title("Cronograma de Formação 2025")
    
    input_file = 'tarefas.xlsx'
    try:
        df = pd.read_excel(input_file)
    except FileNotFoundError:
        st.error(f"Erro: Ficheiro '{input_file}' não encontrado.")
        return

    # Clean column names
    df.columns = df.columns.str.strip()
    
    # --- Feature: Mark as Done ---
    # Ensure 'Concluído' column exists
    if 'Concluído' not in df.columns:
        df['Concluído'] = False
    else:
        # Ensure boolean type for existing column
        df['Concluído'] = df['Concluído'].astype(bool)

    st.write("### Lista de Tarefas (Edite para marcar como 'Concluído')")
    # Allow user to edit "Concluído" column
    edited_df = st.data_editor(
        df, 
        column_config={
            "Concluído": st.column_config.CheckboxColumn(
                "Concluído",
                help="Marque para indicar que a tarefa foi realizada",
                default=False,
            )
        },
        disabled=["Semana Sugestiva", "Duração (horas)", "Tema da Formação", "Formador"],
        hide_index=True,
    )
    
    # Allow saving the changes locally
    import io
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        edited_df.to_excel(writer, index=False)
        
    st.download_button(
        label="💾 Descarregar Excel Atualizado",
        data=buffer,
        file_name="tarefas_atualizadas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Descarregue o ficheiro para guardar as suas alterações no seu computador."
    )

    # Use the edited dataframe for the chart
    df = edited_df

    tasks = []
    
    for index, row in df.iterrows():
        try:
            week = row['Semana Sugestiva']
            duration_hours = row['Duração (horas)']
            topic = row['Tema da Formação']
            trainer = row['Formador']
            
            if pd.isna(week) or pd.isna(topic):
                continue
            
            # Data Cleaning: Fix Week/Month types
            week = int(week)
            topic = str(topic).strip()
            trainer = str(trainer).strip()

            start_date = get_start_date_from_week(week)
            days_needed = max(1, duration_hours / 8) # assume 8h/day
            end_date = start_date + timedelta(days=days_needed)

            # Determine Color based on status
            is_done = row.get('Concluído', False)
            color_group = "Concluído" if is_done else row['Formador']
            opacity_val = 0.5 if is_done else 1.0
            
            # Add status text to label
            status_text = "✅ " if is_done else ""
            label = f"{status_text}{topic} ({trainer})"

            tasks.append(dict(
                Tarefa=topic,
                Formador=trainer,
                Cor=color_group, # Use this new field for coloring
                Início=start_date,
                Fim=end_date,
                Duração=f"{duration_hours}h",
                Semana=f"Semana {week}",
                Label=label,
                Opacity=opacity_val
            ))
            
        except Exception as e:
            st.warning(f"Skipping row {index}: {e}")
            continue

    if not tasks:
        st.warning("Nenhuma tarefa válida encontrada.")
        return

    df_gantt = pd.DataFrame(tasks)
    
    # SORTING IS CRITICAL for "Sequential" look
    # Sort by Start Date then Task Name
    df_gantt = df_gantt.sort_values(by=['Início', 'Tarefa'], ascending=[True, True])

    # Sort by Start Date then Task Name for Waterfall effect
    df_gantt = df_gantt.sort_values(by=['Início', 'Tarefa'], ascending=[True, True])

    # Create Gantt Chart with Dark Template
    fig = px.timeline(
        df_gantt, 
        x_start="Início", 
        x_end="Fim", 
        y="Tarefa", 
        color="Cor", # Use our custom color group
        text="Label",
        opacity=0.9, # Base opacity
        # Define specific colors if needed, otherwise Plotly assigns them
        color_discrete_map={
            "Concluído": "gray" # Force gray for completed tasks
        },
        hover_data={"Duração": True, "Semana": True, "Início": False, "Fim": False, "Label": False},
        title="Cronograma de Formação 2025",
        template='plotly_dark' # DARK MODE
    )

    # REVERSE Y-axis
    fig.update_yaxes(autorange="reversed", title="") 
    
    # Calculate range padding
    min_date = df_gantt['Início'].min() - timedelta(days=2) # Reduced padding
    max_date = df_gantt['Fim'].max() + timedelta(days=2) # Reduced padding

    fig.update_xaxes(
        title="", 
        tickformat="%d %b", 
        dtick="604800000", # Explicit formatting for 7 days in ms
        side="top",
        range=[min_date, max_date],
        showgrid=True,
        gridwidth=1,
        gridcolor='#444' # Subtle grid in dark mode
    )
    
    fig.update_layout(
        font=dict(family="Arial", size=14, color="#eee"),
        title_font_size=26,
        hoverlabel=dict(bgcolor="#333", font_size=14, font_family="Arial"),
        height=500 + (len(df_gantt) * 45), 
        yaxis_showgrid=True,
        yaxis_gridcolor='#444',
        plot_bgcolor='#222', # Dark background
        paper_bgcolor='#222',
        showlegend=True,
        legend_title_text='Formador',
        uniformtext_minsize=10, 
        uniformtext_mode='hide',
        margin=dict(l=20, r=20, t=100, b=50)
    )
    
    # Update bars styling
    fig.update_traces(
        marker_line_color='white', # White border for contrast
        marker_line_width=1, 
        opacity=0.9, 
        textposition='outside', # Place text outside bar to avoid clipping
        cliponaxis=False, # Allow text to overflow plot area
        textfont_size=12
    )

    st.plotly_chart(fig, width="stretch")

if __name__ == "__main__":
    generate_schedule()
