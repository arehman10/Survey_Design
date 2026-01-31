"""
HTML report generation.

Generates interactive HTML reports with styling and tables.
"""

import pandas as pd
from typing import Dict, Any, List, Tuple
import logging
import html

logger = logging.getLogger(__name__)


class HTMLGenerator:
    """Generator for HTML reports."""
    
    HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <style>
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }}
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        h1 {{
            color: #1a1a1a;
            border-bottom: 3px solid #366092;
            padding-bottom: 10px;
        }}
        h2 {{
            color: #366092;
            margin-top: 30px;
            margin-bottom: 15px;
        }}
        h3 {{
            color: #555;
            margin-top: 20px;
        }}
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 15px 0;
            background-color: white;
        }}
        th, td {{
            border: 1px solid #ddd;
            padding: 10px;
            text-align: left;
        }}
        th {{
            background-color: #366092;
            color: white;
            font-weight: 600;
        }}
        tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}
        tr:hover {{
            background-color: #f0f0f0;
        }}
        .numeric {{
            text-align: right;
        }}
        .section {{
            margin: 30px 0;
            padding: 20px;
            background-color: #fafafa;
            border-radius: 4px;
        }}
        .metric-card {{
            display: inline-block;
            margin: 10px;
            padding: 15px 25px;
            background-color: white;
            border-radius: 4px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }}
        .metric-label {{
            font-size: 12px;
            color: #666;
            text-transform: uppercase;
        }}
        .metric-value {{
            font-size: 24px;
            font-weight: 600;
            color: #366092;
        }}
        .scrollbox {{
            max-height: 500px;
            overflow-y: auto;
            border: 1px solid #ddd;
            border-radius: 4px;
        }}
        .timestamp {{
            color: #999;
            font-size: 14px;
            margin-top: 20px;
        }}
        .success {{
            color: #28a745;
        }}
        .warning {{
            color: #ffc107;
            background-color: #fff3cd;
            padding: 10px;
            border-radius: 4px;
            margin: 10px 0;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>{title}</h1>
        {content}
        <div class="timestamp">Generated: {timestamp}</div>
    </div>
</body>
</html>
"""
    
    def __init__(self):
        """Initialize HTML generator."""
        pass
    
    def generate(
        self,
        results: Dict[str, Any],
        input_data: pd.DataFrame
    ) -> bytes:
        """
        Generate HTML report.
        
        Args:
            results: Optimization results
            input_data: Original input data
            
        Returns:
            HTML report as bytes
        """
        from datetime import datetime
        
        try:
            content_sections = []
            
            # Summary section
            content_sections.append(self._create_summary_section(results))
            
            # Scenario sections
            for scenario_name, result in results.items():
                if result.get('success'):
                    content_sections.append(
                        self._create_scenario_section(scenario_name, result)
                    )
            
            # Combine all sections
            content = "\n".join(content_sections)
            
            # Generate HTML
            html_content = self.HTML_TEMPLATE.format(
                title="Sample Allocation Optimization Report",
                content=content,
                timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            )
            
            logger.info("HTML report generated successfully")
            return html_content.encode('utf-8')
            
        except Exception as e:
            logger.error(f"Error generating HTML report: {e}", exc_info=True)
            raise
    
    def _create_summary_section(self, results: Dict[str, Any]) -> str:
        """Create summary section."""
        html_parts = ['<div class="section">']
        html_parts.append('<h2>Summary</h2>')
        
        # Metric cards
        for scenario_name, result in results.items():
            if result.get('success'):
                scenario_label = scenario_name.replace('scenario_', 'Scenario ')
                html_parts.append(f'<h3>{scenario_label}</h3>')
                
                # Key metrics
                metrics = [
                    ('Total Sample', result.get('total_sample', 0)),
                    ('Panel Sample', result.get('panel_sample', 0)),
                    ('Fresh Sample', result.get('fresh_sample', 0)),
                    ('Panel %', f"{result.get('panel_pct', 0):.1f}%"),
                    ('Fresh %', f"{result.get('fresh_pct', 0):.1f}%"),
                ]
                
                for label, value in metrics:
                    html_parts.append(f'''
                    <div class="metric-card">
                        <div class="metric-label">{label}</div>
                        <div class="metric-value">{value:,}</div>
                    </div>
                    ''')
                
                # Warnings if any
                if result.get('warnings'):
                    html_parts.append('<div class="warning">')
                    html_parts.append('<strong>Warnings:</strong><ul>')
                    for warning in result['warnings']:
                        html_parts.append(f'<li>{html.escape(warning)}</li>')
                    html_parts.append('</ul></div>')
        
        html_parts.append('</div>')
        return '\n'.join(html_parts)
    
    def _create_scenario_section(
        self,
        scenario_name: str,
        result: Dict[str, Any]
    ) -> str:
        """Create section for a scenario."""
        html_parts = ['<div class="section">']
        scenario_label = scenario_name.replace('scenario_', 'Scenario ')
        html_parts.append(f'<h2>{scenario_label} - Detailed Results</h2>')
        
        # Region totals
        if 'region_totals' in result:
            html_parts.append('<h3>Region Totals</h3>')
            html_parts.append('<div class="scrollbox">')
            html_parts.append(
                result['region_totals'].to_html(index=False, classes='')
            )
            html_parts.append('</div>')
        
        # Size totals
        if 'size_totals' in result:
            html_parts.append('<h3>Size Totals</h3>')
            html_parts.append('<div class="scrollbox">')
            html_parts.append(
                result['size_totals'].to_html(index=False, classes='')
            )
            html_parts.append('</div>')
        
        # Industry totals
        if 'industry_totals' in result:
            html_parts.append('<h3>Industry Totals</h3>')
            html_parts.append('<div class="scrollbox">')
            html_parts.append(
                result['industry_totals'].to_html(index=False, classes='')
            )
            html_parts.append('</div>')
        
        html_parts.append('</div>')
        return '\n'.join(html_parts)
