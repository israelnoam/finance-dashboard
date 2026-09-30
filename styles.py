"""
Custom CSS styling for the Finance Dashboard.
Themes:
  1. Dark mode (#0d0d12) with emerald green & crimson red glowing accents.
  2. Light mode (#f8fafc) with clean modern cards, soft shadows & vibrant accents.
"""

def get_custom_css(theme: str = "dark") -> str:
    is_light = (theme == "light")
    
    if is_light:
        app_bg = "#f8fafc"
        text_primary = "#0f172a"
        text_muted = "#64748b"
        title_gradient = "linear-gradient(135deg, #0f172a 0%, #334155 100%)"
        header_border = "rgba(0, 0, 0, 0.08)"
        card_bg = "#ffffff"
        card_border = "#e2e8f0"
        card_shadow = "0 2px 8px rgba(0, 0, 0, 0.04)"
        sidebar_bg = "#f1f5f9"
        
        income_color = "#059669"
        income_border = "rgba(5, 150, 105, 0.35)"
        income_bg = "linear-gradient(145deg, #ffffff 0%, #ecfdf5 100%)"
        income_glow = "none"
        
        expense_color = "#dc2626"
        expense_border = "rgba(220, 38, 38, 0.35)"
        expense_bg = "linear-gradient(145deg, #ffffff 0%, #fef2f2 100%)"
        expense_glow = "none"
        
        net_pos_color = "#0284c7"
        net_pos_border = "rgba(2, 132, 199, 0.35)"
        net_pos_bg = "linear-gradient(145deg, #ffffff 0%, #f0f9ff 100%)"
        net_pos_glow = "none"
        
        net_neg_color = "#e11d48"
        net_neg_border = "rgba(225, 29, 72, 0.35)"
        net_neg_bg = "linear-gradient(145deg, #ffffff 0%, #fff1f2 100%)"
        net_neg_glow = "none"
        
        modal_bg = "#ffffff"
        modal_border = "#e2e8f0"
        modal_shadow = "0 20px 60px rgba(0, 0, 0, 0.12)"
    else:
        app_bg = "#0d0d12"
        text_primary = "#f3f4f6"
        text_muted = "#94a3b8"
        title_gradient = "linear-gradient(135deg, #ffffff 0%, #a1a1aa 100%)"
        header_border = "rgba(255, 255, 255, 0.08)"
        card_bg = "#14141e"
        card_border = "#232336"
        card_shadow = "0 4px 16px rgba(0, 0, 0, 0.2)"
        sidebar_bg = "#0d0d12"
        
        income_color = "#10b981"
        income_border = "rgba(16, 185, 129, 0.3)"
        income_bg = "linear-gradient(145deg, #151520 0%, #11111a 100%)"
        income_glow = "0 0 16px rgba(16, 185, 129, 0.4)"
        
        expense_color = "#f43f5e"
        expense_border = "rgba(239, 68, 68, 0.3)"
        expense_bg = "linear-gradient(145deg, #151520 0%, #11111a 100%)"
        expense_glow = "0 0 16px rgba(244, 63, 94, 0.4)"
        
        net_pos_color = "#38bdf8"
        net_pos_border = "rgba(56, 189, 248, 0.3)"
        net_pos_bg = "linear-gradient(145deg, #151520 0%, #11111a 100%)"
        net_pos_glow = "0 0 16px rgba(56, 189, 248, 0.4)"
        
        net_neg_color = "#fb7185"
        net_neg_border = "rgba(244, 63, 94, 0.3)"
        net_neg_bg = "linear-gradient(145deg, #151520 0%, #11111a 100%)"
        net_neg_glow = "0 0 16px rgba(251, 113, 133, 0.4)"
        
        modal_bg = "#12121c"
        modal_border = "#27273a"
        modal_shadow = "0 20px 60px rgba(0, 0, 0, 0.8), 0 0 30px rgba(16, 185, 129, 0.1)"

    return f"""
    <style>
        /* Import Inter / JetBrains Mono font */
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');

        html, body, [class*="css"] {{
            font-family: 'Inter', sans-serif;
        }}

        /* App Background */
        .stApp {{
            background-color: {app_bg} !important;
            color: {text_primary};
        }}

        /* Top header styling */
        .dashboard-header {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 1.2rem 0;
            margin-bottom: 1.5rem;
            border-bottom: 1px solid {header_border};
        }}

        .dashboard-title {{
            font-size: 1.85rem;
            font-weight: 700;
            letter-spacing: -0.5px;
            background: {title_gradient};
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin: 0;
            display: flex;
            align-items: center;
            gap: 0.6rem;
        }}

        .dashboard-subtitle {{
            font-size: 0.88rem;
            color: {text_muted};
            margin-top: 0.25rem;
        }}

        /* KPI Cards */
        .kpi-container {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 1.2rem;
            margin-bottom: 1.8rem;
        }}

        .kpi-card {{
            border-radius: 14px;
            padding: 1.3rem 1.4rem;
            position: relative;
            overflow: hidden;
            transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease;
        }}

        .kpi-card:hover {{
            transform: translateY(-2px);
        }}

        .kpi-income {{
            background: {income_bg};
            border: 1px solid {income_border};
            box-shadow: 0 4px 15px -2px rgba(16, 185, 129, 0.12);
        }}
        .kpi-income:hover {{
            border-color: rgba(16, 185, 129, 0.6);
            box-shadow: 0 6px 24px 0 rgba(16, 185, 129, 0.22);
        }}

        .kpi-expense {{
            background: {expense_bg};
            border: 1px solid {expense_border};
            box-shadow: 0 4px 15px -2px rgba(239, 68, 68, 0.12);
        }}
        .kpi-expense:hover {{
            border-color: rgba(239, 68, 68, 0.6);
            box-shadow: 0 6px 24px 0 rgba(239, 68, 68, 0.22);
        }}

        .kpi-net-positive {{
            background: {net_pos_bg};
            border: 1px solid {net_pos_border};
            box-shadow: 0 4px 15px -2px rgba(56, 189, 248, 0.12);
        }}
        .kpi-net-positive:hover {{
            border-color: rgba(56, 189, 248, 0.6);
            box-shadow: 0 6px 24px 0 rgba(56, 189, 248, 0.22);
        }}

        .kpi-net-negative {{
            background: {net_neg_bg};
            border: 1px solid {net_neg_border};
            box-shadow: 0 4px 15px -2px rgba(244, 63, 94, 0.12);
        }}
        .kpi-net-negative:hover {{
            border-color: rgba(244, 63, 94, 0.6);
            box-shadow: 0 6px 24px 0 rgba(244, 63, 94, 0.22);
        }}

        .kpi-label {{
            font-size: 0.82rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            color: {text_muted};
            margin-bottom: 0.4rem;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }}

        .kpi-bullet {{
            display: inline-block !important;
            width: 9px !important;
            height: 9px !important;
            min-width: 9px !important;
            min-height: 9px !important;
            border-radius: 50% !important;
            margin-right: 6px !important;
            flex-shrink: 0 !important;
        }}

        .kpi-value {{
            font-size: 2rem;
            font-weight: 700;
            font-family: 'JetBrains Mono', monospace;
            line-height: 1.1;
        }}

        .kpi-value.income {{
            color: {income_color};
            text-shadow: {income_glow};
        }}

        .kpi-value.expense {{
            color: {expense_color};
            text-shadow: {expense_glow};
        }}

        .kpi-value.net-pos {{
            color: {net_pos_color};
            text-shadow: {net_pos_glow};
        }}

        .kpi-value.net-neg {{
            color: {net_neg_color};
            text-shadow: {net_neg_glow};
        }}

        .kpi-subtext {{
            font-size: 0.78rem;
            color: {text_muted};
            margin-top: 0.4rem;
        }}

        /* Filter Section Styling */
        .controls-card {{
            background: {card_bg};
            border: 1px solid {card_border};
            border-radius: 12px;
            padding: 1rem 1.2rem 0.8rem 1.2rem;
            margin-bottom: 1.5rem;
            box-shadow: {card_shadow};
        }}

        /* Modal Dialog Customization */
        div[data-modal-container="true"] > div {{
            background-color: {modal_bg} !important;
            border: 1px solid {modal_border} !important;
            border-radius: 16px !important;
            box-shadow: {modal_shadow} !important;
        }}

        /* Custom buttons */
        .stButton > button {{
            border-radius: 9px !important;
            font-weight: 600 !important;
            transition: all 0.2s ease !important;
        }}

        /* Streamlit Data Editor and Table polish */
        [data-testid="stDataFrame"] {{
            border: 1px solid {card_border};
            border-radius: 10px;
            overflow: hidden;
        }}

        /* Chart container */
        .chart-box {{
            background: {card_bg};
            border: 1px solid {card_border};
            border-radius: 14px;
            padding: 1.2rem;
            box-shadow: {card_shadow};
            margin-bottom: 1.5rem;
        }}

        /* Market Insights Cards */
        .insight-card {{
            background: {card_bg};
            border: 1px solid {card_border};
            border-radius: 12px;
            padding: 1.1rem 1.3rem;
            margin-bottom: 1rem;
            box-shadow: {card_shadow};
            transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease;
        }}
        .insight-card:hover {{
            transform: translateY(-2px);
            border-color: rgba(59, 130, 246, 0.45);
        }}
        .insight-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 0.6rem;
        }}
        .insight-badge-fibi {{
            background: rgba(59, 130, 246, 0.15);
            color: #3b82f6;
            border: 1px solid rgba(59, 130, 246, 0.3);
            padding: 2px 8px;
            border-radius: 6px;
            font-size: 0.75rem;
            font-weight: 600;
        }}
        .insight-badge-deals {{
            background: rgba(244, 63, 94, 0.15);
            color: #f43f5e;
            border: 1px solid rgba(244, 63, 94, 0.3);
            padding: 2px 8px;
            border-radius: 6px;
            font-size: 0.75rem;
            font-weight: 600;
        }}
        .insight-badge-hever {{
            background: rgba(168, 85, 247, 0.15);
            color: #a855f7;
            border: 1px solid rgba(168, 85, 247, 0.3);
            padding: 2px 8px;
            border-radius: 6px;
            font-size: 0.75rem;
            font-weight: 600;
        }}
        .insight-date {{
            font-size: 0.78rem;
            color: {text_muted};
        }}
        .insight-body {{
            font-size: 0.92rem;
            line-height: 1.6;
            color: {text_primary};
            white-space: pre-wrap !important;
            word-break: break-word;
            direction: rtl !important;
            text-align: right !important;
            unicode-bidi: embed !important;
            margin: 0.5rem 0;
        }}
        .insight-footer {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 0.5rem;
            margin-top: 0.7rem;
            padding-top: 0.5rem;
            border-top: 1px solid {card_border};
            font-size: 0.78rem;
            color: {text_muted};
        }}

        /* iOS-style Pills component styling */
        div[data-testid="stPills"] {{
            gap: 8px !important;
            flex-wrap: wrap !important;
        }}
        div[data-testid="stPills"] button {{
            border-radius: 20px !important;
            font-size: 0.85rem !important;
            font-weight: 600 !important;
            padding: 6px 14px !important;
            border: 1px solid rgba(255, 255, 255, 0.14) !important;
            transition: all 0.2s ease !important;
        }}
        div[data-testid="stPills"] button[aria-selected="true"] {{
            background: linear-gradient(135deg, #1e40af, #3b82f6) !important;
            color: #ffffff !important;
            border-color: #60a5fa !important;
            box-shadow: 0 2px 10px rgba(37, 99, 235, 0.4) !important;
        }}

        /* Merged Category List Item */
        .cat-item-card {{
            background: rgba(255, 255, 255, 0.03);
            border: 1px solid rgba(255, 255, 255, 0.07);
            border-radius: 10px;
            padding: 0.75rem 1rem;
            margin-bottom: 0.5rem;
            transition: all 0.15s ease;
        }}
        .cat-item-card:hover {{
            background: rgba(59, 130, 246, 0.12);
            border-color: rgba(59, 130, 246, 0.4);
        }}
        .cat-progress-track {{
            height: 6px;
            background: rgba(255, 255, 255, 0.08);
            border-radius: 3px;
            overflow: hidden;
            margin-top: 5px;
        }}
        .cat-progress-bar {{
            height: 100%;
            border-radius: 3px;
        }}

        /* Category Drill-Down Box */
        .category-drilldown-box {{
            background: {card_bg};
            border: 1px solid {card_border};
            border-left: 4px solid #3b82f6;
            border-radius: 12px;
            padding: 1.2rem;
            margin-top: 1rem;
            margin-bottom: 1.5rem;
            box-shadow: {card_shadow};
        }}

        /* Streamlit Category Pills Centered & Touch-Friendly */
        div[data-testid="stPills"] {{
            display: flex !important;
            justify-content: center !important;
            flex-wrap: wrap !important;
            gap: 0.45rem !important;
            margin-top: 0.2rem !important;
            margin-bottom: 0.4rem !important;
        }}
        div[data-testid="stPills"] button {{
            border-radius: 20px !important;
            font-size: 0.82rem !important;
            padding: 0.28rem 0.75rem !important;
            transition: all 0.15s ease-in-out !important;
            display: inline-flex !important;
            align-items: center !important;
            justify-content: center !important;
        }}

        /* Mobile Viewport Optimizations (max-width: 768px) */
        @media screen and (max-width: 768px) {{
            .dashboard-header {{
                flex-direction: column;
                align-items: flex-start;
                padding: 0.8rem 0;
                margin-bottom: 1rem;
            }}
            .dashboard-title {{
                font-size: 1.4rem !important;
            }}
            .kpi-container {{
                grid-template-columns: 1fr !important;
                gap: 0.75rem !important;
                margin-bottom: 1.2rem !important;
            }}
            .kpi-card {{
                padding: 0.9rem 1.1rem !important;
                border-radius: 12px !important;
            }}
            .kpi-value {{
                font-size: 1.6rem !important;
            }}
            .kpi-label {{
                font-size: 0.78rem !important;
            }}
            .controls-card {{
                padding: 0.8rem 0.9rem !important;
                margin-bottom: 1rem !important;
            }}
            .chart-box {{
                padding: 0.9rem !important;
                border-radius: 12px !important;
                margin-bottom: 1rem !important;
            }}
            .insight-card {{
                padding: 0.9rem 1rem !important;
            }}
            .insight-body {{
                font-size: 0.88rem !important;
                line-height: 1.6 !important;
                direction: rtl !important;
                text-align: right !important;
                unicode-bidi: embed !important;
                white-space: pre-wrap !important;
            }}
            /* Ensure buttons and inputs are easily tappable */
            .stButton > button {{
                min-height: 44px !important;
                border-radius: 10px !important;
            }}
            div[data-testid="column"] {{
                min-width: 100% !important;
                margin-bottom: 0.5rem;
            }}
        }}
    </style>
    """
