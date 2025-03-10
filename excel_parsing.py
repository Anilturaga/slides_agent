def get_excel_table(file_path: str, sheet_name: str, max_rows: int = 15) -> str:
    """
    Get the Excel table as a markdown table, with clear schema information and grounding metrics.
    Only describes hierarchical structure when present.
    Provides data type-specific metrics for each column to better understand the data distribution.
    
    Args:
        file_path: Path to the Excel file
        sheet_name: Name of the sheet to read
        max_rows: Maximum number of rows to return (default: 15)
    
    Returns:
        Structured markdown representation with schema, metrics, and sample data
    """
    import pandas as pd
    import numpy as np
    
    try:
        # Read Excel file - read all rows to calculate metrics, but only display max_rows
        df_full = pd.read_excel(file_path, sheet_name=sheet_name)
        df = df_full.head(max_rows)
        
        # Get total row count
        total_rows = len(df_full)
        
        # Check for hierarchical structures
        has_multiindex_columns = isinstance(df.columns, pd.MultiIndex)
        has_multiindex_rows = isinstance(df.index, pd.MultiIndex) and not df.index.names[0] is None
        
        result = []
        
        # Basic information about the table
        result.append(f"## Excel Table: {sheet_name}")
        result.append(f"Total rows: {total_rows} (showing {min(max_rows, len(df))})")
        
        # Schema information and column metrics
        result.append("\n### Schema and Metrics:")
        
        if not has_multiindex_columns:
            # Create schema with column-specific metrics
            schema_data = []
            
            for col in df_full.columns:
                col_info = {'Column': col, 'Data Type': str(df_full[col].dtype)}
                
                # Add type-specific metrics
                if pd.api.types.is_numeric_dtype(df_full[col]):
                    # For numeric columns: min, max, median, % missing
                    col_info['Min'] = f"{df_full[col].min()}"
                    col_info['Max'] = f"{df_full[col].max()}"
                    col_info['Median'] = f"{df_full[col].median()}"
                    missing_pct = (df_full[col].isna().sum() / len(df_full)) * 100
                    col_info['% Missing'] = f"{missing_pct:.1f}%"
                
                elif pd.api.types.is_string_dtype(df_full[col]) or df_full[col].dtype == 'object':
                    # For categorical/string columns: unique values (limited to top 5)
                    unique_values = df_full[col].dropna().unique()
                    unique_count = len(unique_values)
                    
                    if unique_count <= 5:
                        value_list = ", ".join([str(v) for v in unique_values])
                    else:
                        # Show top 5 most frequent values
                        top_values = df_full[col].value_counts().nlargest(5).index.tolist()
                        value_list = ", ".join([str(v) for v in top_values]) + f" (+ {unique_count-5} more)"
                    
                    col_info['Unique Values'] = f"{unique_count}"
                    col_info['Values'] = value_list
                    missing_pct = (df_full[col].isna().sum() / len(df_full)) * 100
                    col_info['% Missing'] = f"{missing_pct:.1f}%"
                
                elif pd.api.types.is_datetime64_any_dtype(df_full[col]):
                    # For datetime columns: min date, max date, % missing
                    col_info['Min Date'] = str(df_full[col].min())
                    col_info['Max Date'] = str(df_full[col].max())
                    missing_pct = (df_full[col].isna().sum() / len(df_full)) * 100
                    col_info['% Missing'] = f"{missing_pct:.1f}%"
                
                else:
                    # For other types: % missing
                    missing_pct = (df_full[col].isna().sum() / len(df_full)) * 100
                    col_info['% Missing'] = f"{missing_pct:.1f}%"
                
                schema_data.append(col_info)
            
            # Create schema dataframe with all metrics
            schema_df = pd.DataFrame(schema_data)
            result.append(schema_df.to_markdown(index=False))
        
        # Only add hierarchy information if it exists
        if has_multiindex_columns or has_multiindex_rows:
            result.append("\n### Hierarchical Structure:")
            
            if has_multiindex_columns:
                # Show column hierarchy
                levels = df.columns.names
                if all(x is None for x in levels):
                    levels = [f"Level {i}" for i in range(len(df.columns.levels))]
                result.append(f"Column hierarchy: {' > '.join([str(l) for l in levels if l is not None])}")
                
                # Create a visual representation of the column hierarchy
                header_levels = []
                for level in range(df.columns.nlevels):
                    header_levels.append(" | ".join(str(x) for x in df.columns.get_level_values(level)))
                
                result.append("\nColumn header levels:")
                result.append("```")
                result.append("\n".join(header_levels))
                result.append("```")
                
                # Flatten multiindex columns for the markdown table
                df_display = df.copy()
                df_display.columns = [' → '.join([str(x) for x in col if pd.notna(x)]) for col in df.columns.values]
                
                # Add basic metrics for hierarchical tables
                result.append("\nColumn Metrics Summary:")
                metrics = []
                for col in df_full.columns:
                    flat_col_name = ' → '.join([str(x) for x in col if pd.notna(x)])
                    col_data = df_full[col]
                    
                    if pd.api.types.is_numeric_dtype(col_data):
                        metrics.append(f"- **{flat_col_name}**: Range [{col_data.min()} to {col_data.max()}], Median: {col_data.median()}")
                    elif pd.api.types.is_string_dtype(col_data) or col_data.dtype == 'object':
                        unique_count = len(col_data.dropna().unique())
                        metrics.append(f"- **{flat_col_name}**: {unique_count} unique values")
                
                result.append("\n".join(metrics))
            else:
                df_display = df.copy()
            
            if has_multiindex_rows:
                # Show row hierarchy
                levels = df.index.names
                if all(x is None for x in levels):
                    levels = [f"Level {i}" for i in range(len(df.index.levels))]
                result.append(f"Row hierarchy: {' > '.join([str(l) for l in levels if l is not None])}")
                
                # Reset index to convert hierarchical rows to columns
                df_display = df_display.reset_index()
        else:
            df_display = df
        
        # Sample data
        result.append("\n### Sample Data:")
        result.append(df_display.to_markdown(index=False))
        
        return "\n".join(result)
    except Exception as e:
        return f"Error processing Excel file: {str(e)}"
