

import pandas as pd
import plotly.graph_objects as go
from collections import defaultdict
import numpy as np

def load_relationships_from_csv(csv_file):
    """Load relationship data from CSV file"""
    df = pd.read_csv(csv_file)
    return df

def create_interactive_er_diagram_from_df(df, output_file,image_format="png"):
    """
    Wrapper to generate ER diagram directly from a DataFrame
    (No CSV written to disk)
    """
    # Reuse the existing logic by passing df directly
    return _create_er_diagram(df, output_file,image_format)

def create_interactive_er_diagram(csv_file, output_file):
    """
    Create an interactive ER diagram using Plotly
    
    Parameters:
    csv_file: Path to CSV file with columns: Table name, Column name, Connects to Table, Connects to column
    output_file: Path to save the HTML output
    """
    # Load data
    df = load_relationships_from_csv(csv_file)
    return _create_er_diagram(df, output_file)

def _create_er_diagram(df, output_file,image_format="png"):
    """
    Core ER diagram logic (shared by CSV & DataFrame versions)
    """

    # 🔽 PASTE ALL YOUR EXISTING PLOTLY CODE HERE 🔽
    # starting from:
    #   tables = defaultdict(set)
    # until:
    #   fig.write_html(output_file)
    #   return fig
    # Extract unique tables and their columns
    tables = defaultdict(set)
    for _, row in df.iterrows():
        tables[row['Table name']].add(row['Column name'])
        tables[row['Connects to Table']].add(row['Connects to column'])
    
    # Convert to sorted lists
    tables = {table: sorted(list(cols)) for table, cols in tables.items()}
    
    # Separate tables into two groups
    fk_tables = set(df['Table name'].unique())
    pk_tables = set(df['Connects to Table'].unique())
    
    left_tables = sorted(list(fk_tables))
    right_tables = sorted(list(pk_tables))
    
    print(f"\n📊 Creating Interactive ER Diagram:")
    print(f"Left side (Foreign Key tables): {left_tables}")
    print(f"Right side (Primary Key tables): {right_tables}")
    
    # Layout parameters
    table_width = 250
    table_height_base = 40
    column_height = 25
    vertical_spacing = 80
    horizontal_gap = 400
    left_x = 50
    right_x = left_x + table_width + horizontal_gap
    
    # Calculate positions and store table info
    table_positions = {}
    
    # Position left tables - Y coordinate starts at TOP
    current_y = 50
    for table_name in left_tables:
        columns = tables.get(table_name, [])
        table_height = table_height_base + len(columns) * column_height
        
        table_positions[table_name] = {
            'x': left_x,
            'y': current_y,  # This is the TOP of the table (header starts here)
            'width': table_width,
            'height': table_height,
            'columns': {},
            'side': 'left',
            'is_primary': table_name in pk_tables
        }
        
        # Store column positions - columns go DOWNWARD from header
        for i, col_name in enumerate(columns):
            # Columns start AFTER the header (y + table_height_base) and go DOWN
            col_y = current_y + table_height_base + (i + 0.5) * column_height
            table_positions[table_name]['columns'][col_name] = col_y
        
        # Next table goes BELOW this one
        current_y += table_height + vertical_spacing
    
    # Position right tables - Y coordinate starts at TOP
    current_y = 50
    for table_name in right_tables:
        if table_name in table_positions:
            continue  # Already positioned on left
            
        columns = tables.get(table_name, [])
        table_height = table_height_base + len(columns) * column_height
        
        table_positions[table_name] = {
            'x': right_x,
            'y': current_y,  # This is the TOP of the table (header starts here)
            'width': table_width,
            'height': table_height,
            'columns': {},
            'side': 'right',
            'is_primary': True
        }
        
        # Store column positions - columns go DOWNWARD from header
        for i, col_name in enumerate(columns):
            # Columns start AFTER the header (y + table_height_base) and go DOWN
            col_y = current_y + table_height_base + (i + 0.5) * column_height
            table_positions[table_name]['columns'][col_name] = col_y
        
        # Next table goes BELOW this one
        current_y += table_height + vertical_spacing
    
    # Create Plotly figure
    fig = go.Figure()
    
    # Draw tables
    for table_name, pos in table_positions.items():
        columns = tables.get(table_name, [])
        
        # Table colors
        header_color = "#E97E12" if pos['is_primary'] else '#E97E12'
        
        # Draw table header at the TOP
        fig.add_shape(
            type="rect",
            x0=pos['x'], y0=pos['y'],
            x1=pos['x'] + pos['width'], y1=pos['y'] + table_height_base,
            fillcolor=header_color,
            line=dict(color='black', width=2),
            layer='above'
        )
        
        # Add table name at the TOP
        fig.add_annotation(
            x=pos['x'] + pos['width']/2,
            y=pos['y'] + table_height_base/2,
            text=f"<b>{table_name}</b>",
            showarrow=False,
            font=dict(size=12, color='white'),
            xanchor='center',
            yanchor='middle'
        )
        
        # Draw table body BELOW the header
        fig.add_shape(
            type="rect",
            x0=pos['x'], y0=pos['y'] + table_height_base,
            x1=pos['x'] + pos['width'], y1=pos['y'] + pos['height'],
            fillcolor='white',
            line=dict(color='black', width=1.5),
            layer='above'
        )
        
        # Add columns BELOW the header
        for i, col_name in enumerate(columns):
            col_y = pos['y'] + table_height_base + (i + 0.5) * column_height
            
            # Check if PK or FK
            is_pk = any((df['Connects to Table'] == table_name) & 
                       (df['Connects to column'] == col_name))
            is_fk = any((df['Table name'] == table_name) & 
                       (df['Column name'] == col_name))
            
            if is_pk:
                icon = "🔑 "
                color = '#D84315'
            elif is_fk:
                icon = "🔗 "
                color = '#1976D2'
            else:
                icon = "   "
                color = 'black'
            
            display_name = icon + col_name
            if len(display_name) > 30:
                display_name = display_name[:27] + '...'
            
            fig.add_annotation(
                x=pos['x'] + 10,
                y=col_y,
                text=display_name,
                showarrow=False,
                font=dict(size=10, color=color),
                xanchor='left',
                yanchor='middle'
            )
    
    # Define distinct colors for different table pairs (avoiding white and very light colors)
    color_palette = [
        '#E91E63',  # Pink
        '#9C27B0',  # Purple
        '#3F51B5',  # Indigo
        '#2196F3',  # Blue
        '#00BCD4',  # Cyan
        '#009688',  # Teal
        '#4CAF50',  # Green
        '#FF9800',  # Orange
        '#FF5722',  # Deep Orange
        '#795548',  # Brown
        '#607D8B',  # Blue Grey
        '#F44336',  # Red
        '#673AB7',  # Deep Purple
        '#8BC34A',  # Light Green
        '#FFC107',  # Amber
        '#00897B',  # Teal Dark
        '#5E35B1',  # Deep Purple Dark
        '#D32F2F',  # Red Dark
    ]
    
    # Assign colors to unique table pairs
    table_pair_colors = {}
    color_index = 0
    
    # First pass: identify unique table pairs and assign colors
    for idx, row in df.iterrows():
        from_table = row['Table name']
        to_table = row['Connects to Table']
        
        # Create a canonical pair key (sorted to treat A->B and B->A as same pair)
        pair_key = tuple(sorted([from_table, to_table]))
        
        if pair_key not in table_pair_colors:
            table_pair_colors[pair_key] = color_palette[color_index % len(color_palette)]
            color_index += 1
    
    print(f"\n🎨 Color assignments for table pairs:")
    for pair, color in table_pair_colors.items():
        print(f"  {pair[0]} ↔ {pair[1]}: {color}")
    
    # Draw relationships - each as a separate straight line with color by table pair
    print(f"\n🔗 Drawing {len(df)} relationships...")
    
    # Track how many relationships have been drawn to each column to add offset
    column_relationship_count = defaultdict(int)
    
    for idx, row in df.iterrows():
        from_table = row['Table name']
        from_col = row['Column name']
        to_table = row['Connects to Table']
        to_col = row['Connects to column']
        
        print(f"  → {from_table}.{from_col} -> {to_table}.{to_col}")
        
        if from_table not in table_positions or to_table not in table_positions:
            continue
        
        from_pos = table_positions[from_table]
        to_pos = table_positions[to_table]
        
        if from_col not in from_pos['columns'] or to_col not in to_pos['columns']:
            continue
        
        # Get the color for this table pair
        pair_key = tuple(sorted([from_table, to_table]))
        line_color = table_pair_colors[pair_key]
        
        # Get column positions
        from_y = from_pos['columns'][from_col]
        to_y = to_pos['columns'][to_col]
        
        # Create unique key for this connection pair
        connection_key = f"{from_table}.{from_col}-{to_table}.{to_col}"
        offset_index = column_relationship_count[connection_key]
        column_relationship_count[connection_key] += 1
        
        # Add small vertical offset if multiple relationships use same columns
        y_offset = offset_index * 5
        
        # Connection points with offset
        from_x = from_pos['x'] + from_pos['width']
        to_x = to_pos['x']
        from_y_adj = from_y + y_offset
        to_y_adj = to_y + y_offset
        
        # Draw STRAIGHT line from source to target with table-pair color
        path_x = [from_x, to_x]
        path_y = [from_y_adj, to_y_adj]
        
        # Hover text
        hover_text = f"{from_table}.{from_col} → {to_table}.{to_col}"
        
        # Draw line as separate trace with unique color
        fig.add_trace(go.Scatter(
            x=path_x,
            y=path_y,
            mode='lines',
            line=dict(color=line_color, width=3),
            hovertext=hover_text,
            hoverinfo='text',
            showlegend=False,
            name=''
        ))
        
        # Add arrow at the end (same color as line)
        arrow_length = 15
        # Calculate arrow direction
        dx = to_x - from_x
        dy = to_y_adj - from_y_adj
        length = np.sqrt(dx**2 + dy**2)
        
        if length > 0:
            # Normalize direction
            dx_norm = dx / length
            dy_norm = dy / length
            
            # Arrow start point
            arrow_start_x = to_x - arrow_length * dx_norm
            arrow_start_y = to_y_adj - arrow_length * dy_norm
            
            fig.add_annotation(
                x=to_x,
                y=to_y_adj,
                ax=arrow_start_x,
                ay=arrow_start_y,
                xref='x',
                yref='y',
                axref='x',
                ayref='y',
                showarrow=True,
                arrowhead=2,
                arrowsize=1.5,
                arrowwidth=3,
                arrowcolor=line_color
            )
        
        # Add cardinality labels with color matching the line
        # Calculate position along the line for labels
        label_offset = 20
        
        # N label near source
        fig.add_annotation(
            x=from_x + label_offset,
            y=from_y_adj,
            text='<b>N</b>',
            showarrow=False,
            font=dict(size=10, color='white'),
            bgcolor=line_color,
            bordercolor=line_color,
            borderwidth=2,
            borderpad=3
        )
        
        # 1 label near target
        fig.add_annotation(
            x=to_x - label_offset,
            y=to_y_adj,
            text='<b>1</b>',
            showarrow=False,
            font=dict(size=10, color='white'),
            bgcolor=line_color,
            bordercolor=line_color,
            borderwidth=2,
            borderpad=3
        )
    
    # Calculate figure dimensions
    max_y = max(pos['y'] + pos['height'] for pos in table_positions.values())
    
    # Update layout
    fig.update_layout(
        title=dict(
            text='<b>Interactive Entity-Relationship Diagram</b>',
            font=dict(size=20),
            x=0.5,
            xanchor='center'
        ),
        width=1200,
        height=max(800, max_y + 100),
        xaxis=dict(
            range=[-50, right_x + table_width + 50],
            showgrid=False,
            zeroline=False,
            showticklabels=False
        ),
        yaxis=dict(
            range=[max_y + 50, -50],
            showgrid=False,
            zeroline=False,
            showticklabels=False,
            scaleanchor='x',
            scaleratio=1
        ),
        plot_bgcolor='white',
        hovermode='closest',
        showlegend=True
    )
    
    # Add legend
    fig.add_trace(go.Scatter(
        x=[None], y=[None],
        mode='markers',
        marker=dict(size=10, color='#4CAF50', symbol='square'),
        name='Primary Tables (Referenced)',
        showlegend=True
    ))
    
    fig.add_trace(go.Scatter(
        x=[None], y=[None],
        mode='markers',
        marker=dict(size=10, color='#2196F3', symbol='square'),
        name='Foreign Key Tables',
        showlegend=True
    ))
    
    fig.add_trace(go.Scatter(
        x=[None], y=[None],
        mode='lines',
        line=dict(color='#424242', width=2),
        name='1:N Relationship',
        showlegend=True
    ))
    
    # Save as HTML
    fig.write_html(output_file)
    image_path = output_file.replace(".html", f".{image_format}")

    fig.write_image(
        image_path,
        width=1600,
        height=1000,
        scale=2
    )
    print(f"\n✅ Interactive ER Diagram saved to {output_file}")
    print(f"📊 Created diagram with {len(tables)} tables and {len(df)} relationships")
    print(f"\n🎯 Interactive Features:")
    print("  ✓ Zoom and pan")
    print("  ✓ Hover over relationships to see details")
    print("  ✓ Click and drag to explore")
    print("  ✓ Export as PNG from the toolbar")
    print(f"🖼 ER Diagram image saved to: {image_path}")
    return image_path

# Usage example
if __name__ == "__main__":
    # Create the interactive diagram
    csv_file = 'Doc_op\Relationship.csv'
    fig = create_interactive_er_diagram(csv_file, 'er_diagram_interactive7.html')
    
    print("\n💡 Open 'er_diagram_interactive.html' in your browser to interact!")
    print("\n🚀 Plotly Features:")
    print("  ✓ Fully interactive with zoom/pan")
    print("  ✓ Hover tooltips showing relationship details")
    print("  ✓ Export to PNG/SVG from the modebar")
    print("  ✓ Can be embedded in web apps")
    print("  ✓ Responsive and professional looking") 