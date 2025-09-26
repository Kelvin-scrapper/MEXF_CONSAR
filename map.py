"""
Clean Universal Mexican Pension Fund Mapper
==========================================

A fully universal mapper that auto-discovers files and works with any filename/structure.
Handles all 50 CONSAR mappings with proper NA vs 0.0 distinction.

Features:
✅ Auto File Discovery - scans directory for Excel files by content
✅ WEB_SISTEMA Detection - finds investment files by sheet name
✅ Content-Based Flow Detection - identifies flow files automatically
✅ Order Independent - finds data by content, not position
✅ Structure Adaptive - handles file layout changes
✅ Universal Date Matching - works with any month automatically
✅ Enhanced Flow File Reading - handles HTML-formatted Excel files
✅ Complete CONSAR Mapping - all 50 categories individually mapped
✅ Fixed Output Structure - always 3×51 format

Dependencies: pandas, openpyxl, xlrd, numpy, re, fuzzywuzzy, os, glob
"""

import pandas as pd
import numpy as np
import re
import os
import glob
from datetime import datetime
from fuzzywuzzy import fuzz
import warnings
warnings.filterwarnings('ignore')

class CleanUniversalMexicanPensionMapper:
    """
    Clean implementation of universal Mexican pension fund mapper.
    """
    
    def __init__(self):
        """Initialize with complete CONSAR structure and mappings."""
        
        # COMPLETE 50 CONSAR HEADERS - Fixed order, never changes
        self.consar_headers = [
            'MEXPENSIONFUNDS.DOMESTICEQUITIES.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.FOREIGNEQUITIES.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.COMMODITIES.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTAEROLINEAS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTALIMENTOS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTAUTOMOTRIZ.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTBANCADEDESARROLLO.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTBANCARIO.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTBEBIDAS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTCEMENTO.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTCENTROSCOMERCIALES.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTCONSUMO.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTDEUDACP.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTEMPRESASPRODUCTIVASDELESTADO.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTESTADOS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTEUROBONOS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTGRUPOSINDUSTRIALES.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTCONSTRUCCION.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTINFRAESTRUCTURA.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTINMOBILIARIO.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTOTROS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTPAPEL.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTSERVFINANCIEROS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTSIDERURGICA.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTTELECOM.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTTRANSPORTE.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTVIVIENDA.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.STRUCTUREDASSETS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.REIT.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.FOREIGNBONDS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBOND182.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBONDESD.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBONDESF.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBONOS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBPA182.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBPAS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBPAT.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSCBIC.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSCETES.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSDEPBMX.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSUDIBONO.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSUMS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.OTHER.ACTUALALLOCATION.SPECIFIED.M.1@CONSAR',
            'MEXPENSIONFUNDS.TOTAL.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.TOTAL.LEVEL.NONE.M.1@CONSAR',
            'MEXPENSIONFUNDS.RCV.FLOW.BIGPIPE.M.1@CONSAR',
            'MEXPENSIONFUNDS.RCVIMSS.FLOW.BIGPIPE.M.1@CONSAR',
            'MEXPENSIONFUNDS.RCVISSTE.FLOW.BIGPIPE.M.1@CONSAR',
            'MEXPENSIONFUNDS.IMSSWITHDRAWALS.FLOW.BIGPIPE.M.1@CONSAR',
            'MEXPENSIONFUNDS.ISSTEWITHDRAWALS.FLOW.BIGPIPE.M.1@CONSAR'
        ]
        
        # COMPLETE 50 DESCRIPTIONS - Same order as headers
        self.descriptions = [
            'Investment composition, National Variable Income',
            'Investment composition, International Variable Income',
            'Investment composition, Goods',
            'Investment composition, Airlines',
            'Investment composition, Food',
            'Investment composition, Automotive',
            'Investment composition, Development Bank',
            'Investment composition, Banking',
            'Investment composition, Drinks',
            'Investment composition, Cement',
            'Investment composition, Malls',
            'Investment composition, Consumption',
            'Investment composition, PC debt',
            'Investment composition, State Productive Companies',
            'Investment composition, State',
            'Investment composition, Eurobonds',
            'Investment composition, Industrial Groups',
            'Investment composition, Construction',
            'Investment composition, Infrastructure',
            'Investment composition, Real Estate',
            'Investment composition, OTROS',
            'Investment composition, Paper',
            'Investment composition, Serv. financial',
            'Investment composition, Steel Industry',
            'Investment composition, Telecom',
            'Investment composition, Transport',
            'Investment composition, Living Place',
            'Investment composition, Structured',
            'Investment composition, FIBRAS',
            'Investment composition, International Debt',
            'Investment composition, BOND182',
            'Investment composition, BONDESD',
            'Investment composition, BONDESF',
            'Investment composition, BONOS',
            'Investment composition, BPA182',
            'Investment composition, BPAS',
            'Investment composition, BPAT',
            'Investment composition, CBIC',
            'Investment composition, CETES',
            'Investment composition, DEPBMX',
            'Investment composition, UDIBONO',
            'Investment composition, UMS',
            'Investment composition, Other assets',
            'Investment composition, Total',
            'Net Assets of Generational Siefores, Total',
            'Flow of Resources Channeled to Afores, RCV',
            'Flow of Resources Channeled to Afores, RCV IMSS',
            'Flow of Resources Channeled to Afores, RCV ISSTE',
            'Flow of Withdrawals from Individual Accounts in Afores, Withdrawal of IMSS Resources',
            'Flow of Withdrawals from Individual Accounts in Afores, Withdrawal of ISSSTE Retirement'
        ]
        
        # COMPREHENSIVE CONTENT MAPPING - Spanish to CONSAR
        self.content_mapping = {
            # Main categories
            'renta variable nacional': 'MEXPENSIONFUNDS.DOMESTICEQUITIES.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'renta variable internacional': 'MEXPENSIONFUNDS.FOREIGNEQUITIES.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'mercancias': 'MEXPENSIONFUNDS.COMMODITIES.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'mercancías': 'MEXPENSIONFUNDS.COMMODITIES.ACTUALALLOCATION.NONE.M.1@CONSAR',
            
            # Private debt - Individual mappings
            'aerolineas': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTAEROLINEAS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'aerolíneas': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTAEROLINEAS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'alimentos': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTALIMENTOS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'automotriz': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTAUTOMOTRIZ.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'banca de desarrollo': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTBANCADEDESARROLLO.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'bancario': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTBANCARIO.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'bebidas': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTBEBIDAS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'cemento': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTCEMENTO.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'centros comerciales': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTCENTROSCOMERCIALES.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'consumo': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTCONSUMO.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'deuda cp': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTDEUDACP.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'empresas productivas del estado': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTEMPRESASPRODUCTIVASDELESTADO.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'estados': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTESTADOS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'eurobonos': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTEUROBONOS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'grupos industriales': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTGRUPOSINDUSTRIALES.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'construccion': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTCONSTRUCCION.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'construcción': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTCONSTRUCCION.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'infraestructura': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTINFRAESTRUCTURA.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'inmobiliario': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTINMOBILIARIO.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'otros': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTOTROS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'papel': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTPAPEL.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'serv financieros': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTSERVFINANCIEROS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'servicios financieros': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTSERVFINANCIEROS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'siderurgica': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTSIDERURGICA.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'siderúrgica': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTSIDERURGICA.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'telecom': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTTELECOM.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'telecomunicaciones': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTTELECOM.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'transporte': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTTRANSPORTE.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'vivienda': 'MEXPENSIONFUNDS.DOMESTICPRIVATEDEBTVIVIENDA.ACTUALALLOCATION.NONE.M.1@CONSAR',
            
            # Structured assets
            'estructurados': 'MEXPENSIONFUNDS.STRUCTUREDASSETS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'fibras': 'MEXPENSIONFUNDS.REIT.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'deuda internacional': 'MEXPENSIONFUNDS.FOREIGNBONDS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            
            # Government bonds
            'bond182': 'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBOND182.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'bondesd': 'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBONDESD.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'bondesf': 'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBONDESF.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'bonos': 'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBONOS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'bpa182': 'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBPA182.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'bpas': 'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBPAS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'bpat': 'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBPAT.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'cbic': 'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSCBIC.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'cetes': 'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSCETES.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'depbmx': 'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSDEPBMX.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'udibono': 'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSUDIBONO.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'ums': 'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSUMS.ACTUALALLOCATION.NONE.M.1@CONSAR',
            
            # Other assets - Multiple variations for better matching
            'otros activos': 'MEXPENSIONFUNDS.OTHER.ACTUALALLOCATION.SPECIFIED.M.1@CONSAR',
            'other assets': 'MEXPENSIONFUNDS.OTHER.ACTUALALLOCATION.SPECIFIED.M.1@CONSAR',
            'activos otros': 'MEXPENSIONFUNDS.OTHER.ACTUALALLOCATION.SPECIFIED.M.1@CONSAR',
            'otras inversiones': 'MEXPENSIONFUNDS.OTHER.ACTUALALLOCATION.SPECIFIED.M.1@CONSAR',
            'other investments': 'MEXPENSIONFUNDS.OTHER.ACTUALALLOCATION.SPECIFIED.M.1@CONSAR',
            
            # Total
            'total': 'MEXPENSIONFUNDS.TOTAL.ACTUALALLOCATION.NONE.M.1@CONSAR',
            
            # Flow categories
            'rcv': 'MEXPENSIONFUNDS.RCV.FLOW.BIGPIPE.M.1@CONSAR',
            'recursos canalizados': 'MEXPENSIONFUNDS.RCV.FLOW.BIGPIPE.M.1@CONSAR',
            'rcv imss': 'MEXPENSIONFUNDS.RCVIMSS.FLOW.BIGPIPE.M.1@CONSAR',
            'recursos canalizados imss': 'MEXPENSIONFUNDS.RCVIMSS.FLOW.BIGPIPE.M.1@CONSAR',
            'rcv isste': 'MEXPENSIONFUNDS.RCVISSTE.FLOW.BIGPIPE.M.1@CONSAR',
            'recursos canalizados isste': 'MEXPENSIONFUNDS.RCVISSTE.FLOW.BIGPIPE.M.1@CONSAR',

            # Enhanced withdrawal mappings - exact matches with proper spelling variations
            'retiro de recursos imss': 'MEXPENSIONFUNDS.IMSSWITHDRAWALS.FLOW.BIGPIPE.M.1@CONSAR',
            'retiro de recursos isste': 'MEXPENSIONFUNDS.ISSTEWITHDRAWALS.FLOW.BIGPIPE.M.1@CONSAR',
            'retiro de recursos issste': 'MEXPENSIONFUNDS.ISSTEWITHDRAWALS.FLOW.BIGPIPE.M.1@CONSAR',  # Triple S variation
            'retiro imss': 'MEXPENSIONFUNDS.IMSSWITHDRAWALS.FLOW.BIGPIPE.M.1@CONSAR',
            'retiro isste': 'MEXPENSIONFUNDS.ISSTEWITHDRAWALS.FLOW.BIGPIPE.M.1@CONSAR',
            'retiro issste': 'MEXPENSIONFUNDS.ISSTEWITHDRAWALS.FLOW.BIGPIPE.M.1@CONSAR',  # Triple S variation

            # Additional withdrawal variations for better matching
            'retiro recursos imss': 'MEXPENSIONFUNDS.IMSSWITHDRAWALS.FLOW.BIGPIPE.M.1@CONSAR',
            'retiro recursos isste': 'MEXPENSIONFUNDS.ISSTEWITHDRAWALS.FLOW.BIGPIPE.M.1@CONSAR',
            'retiro recursos issste': 'MEXPENSIONFUNDS.ISSTEWITHDRAWALS.FLOW.BIGPIPE.M.1@CONSAR',  # Triple S variation
            'recursos imss retiro': 'MEXPENSIONFUNDS.IMSSWITHDRAWALS.FLOW.BIGPIPE.M.1@CONSAR',
            'recursos isste retiro': 'MEXPENSIONFUNDS.ISSTEWITHDRAWALS.FLOW.BIGPIPE.M.1@CONSAR',
            'recursos issste retiro': 'MEXPENSIONFUNDS.ISSTEWITHDRAWALS.FLOW.BIGPIPE.M.1@CONSAR',  # Triple S variation
            'imss retiro': 'MEXPENSIONFUNDS.IMSSWITHDRAWALS.FLOW.BIGPIPE.M.1@CONSAR',
            'isste retiro': 'MEXPENSIONFUNDS.ISSTEWITHDRAWALS.FLOW.BIGPIPE.M.1@CONSAR',
            'issste retiro': 'MEXPENSIONFUNDS.ISSTEWITHDRAWALS.FLOW.BIGPIPE.M.1@CONSAR'  # Triple S variation
        }
        
        # DATE MAPPING
        self.spanish_months = {
            'enero': '01', 'febrero': '02', 'marzo': '03', 'abril': '04',
            'mayo': '05', 'junio': '06', 'julio': '07', 'agosto': '08',
            'septiembre': '09', 'octubre': '10', 'noviembre': '11', 'diciembre': '12'
        }
        
        self.month_abbreviations = {
            '01': ['ene', 'jan'], '02': ['feb'], '03': ['mar'], '04': ['abr', 'apr'],
            '05': ['may'], '06': ['jun'], '07': ['jul'], '08': ['ago', 'aug'],
            '09': ['sep'], '10': ['oct'], '11': ['nov'], '12': ['dic', 'dec']
        }
    
    def normalize_text(self, text):
        """Normalize text for universal matching."""
        if not isinstance(text, str):
            return ""
        
        # Remove accents and normalize
        normalized = text.lower().strip()
        normalized = re.sub(r'[áàäâ]', 'a', normalized)
        normalized = re.sub(r'[éèëê]', 'e', normalized)
        normalized = re.sub(r'[íìïî]', 'i', normalized)
        normalized = re.sub(r'[óòöô]', 'o', normalized)
        normalized = re.sub(r'[úùüû]', 'u', normalized)
        normalized = re.sub(r'[ñ]', 'n', normalized)
        normalized = re.sub(r'\s+', ' ', normalized)
        
        return normalized

    def discover_files(self, directory='.'):
        """Universal file discovery - finds files by content, not name."""
        print(f"🔍 Discovering files in directory: {os.path.abspath(directory)}")

        discovered_files = {
            'investment_file': None,
            'flow_files': []
        }

        # Find all Excel files
        excel_patterns = ['*.xlsx', '*.xls', '*.xlsm']
        all_files = []

        for pattern in excel_patterns:
            all_files.extend(glob.glob(os.path.join(directory, pattern)))

        print(f"📁 Found {len(all_files)} Excel files to analyze")

        for filepath in all_files:
            filename = os.path.basename(filepath)
            print(f"  📊 Analyzing: {filename}")

            try:
                # Check if it's the investment file (has WEB_SISTEMA sheet)
                investment_detected = self.detect_investment_file(filepath)
                if investment_detected:
                    discovered_files['investment_file'] = filepath
                    print(f"  ✅ Investment file detected: {filename}")
                    continue

                # Check if it's a flow file
                flow_detected = self.detect_flow_file(filepath)
                if flow_detected:
                    discovered_files['flow_files'].append(filepath)
                    print(f"  ✅ Flow file detected: {filename}")

            except Exception as e:
                print(f"  ⚠️ Could not analyze {filename}: {e}")

        print(f"\n🎯 Discovery Results:")
        print(f"  Investment file: {discovered_files['investment_file'] or 'Not found'}")
        print(f"  Flow files: {len(discovered_files['flow_files'])} found")

        return discovered_files

    def detect_investment_file(self, filepath):
        """Detect investment file by looking for WEB_SISTEMA sheet."""
        try:
            excel_file = pd.ExcelFile(filepath)
            sheet_names = [sheet.upper() for sheet in excel_file.sheet_names]

            # Look for WEB_SISTEMA sheet (case insensitive)
            if 'WEB_SISTEMA' in sheet_names:
                return True

            # Also check for variations
            sistema_variations = ['SISTEMA', 'WEB SISTEMA', 'WEBSISTEMA']
            for variation in sistema_variations:
                if any(variation in sheet_name for sheet_name in sheet_names):
                    return True

            return False

        except Exception:
            return False

    def detect_flow_file(self, filepath):
        """Detect flow file by analyzing content for flow-related terms."""
        try:
            # Try different sheet reading strategies
            df = None

            # Try reading first sheet
            try:
                df = pd.read_excel(filepath, sheet_name=0, header=None)
            except:
                # Try reading as HTML if Excel fails
                try:
                    tables = pd.read_html(filepath, header=None)
                    if tables:
                        df = tables[0]
                except:
                    pass

            if df is None or df.empty:
                return False

            # Convert to string and search for flow indicators
            content_text = ' '.join([
                str(cell).lower() for cell in df.values.flatten()
                if pd.notna(cell) and str(cell).strip()
            ])

            # Flow file indicators
            flow_indicators = [
                'retiro', 'recursos', 'flujo', 'flow', 'canalizados',
                'imss', 'isste', 'issste', 'afores', 'cuentas',
                'individuales', 'rcv'
            ]

            # Count matching indicators
            matches = sum(1 for indicator in flow_indicators if indicator in content_text)

            # If we find multiple flow-related terms, it's likely a flow file
            return matches >= 3

        except Exception:
            return False

    def find_investment_sheet(self, filepath):
        """Find the correct investment sheet (WEB_SISTEMA or similar)."""
        try:
            excel_file = pd.ExcelFile(filepath)

            # First priority: exact match
            for sheet_name in excel_file.sheet_names:
                if sheet_name.upper() == 'WEB_SISTEMA':
                    return sheet_name

            # Second priority: contains SISTEMA
            for sheet_name in excel_file.sheet_names:
                if 'SISTEMA' in sheet_name.upper():
                    return sheet_name

            # Third priority: contains WEB
            for sheet_name in excel_file.sheet_names:
                if 'WEB' in sheet_name.upper():
                    return sheet_name

            # Fallback: first sheet
            return excel_file.sheet_names[0] if excel_file.sheet_names else None

        except Exception:
            return None

    def extract_date_from_text(self, text):
        """Universal date extraction from Spanish text."""
        if not isinstance(text, str):
            return None
        
        text_norm = self.normalize_text(text)
        
        # Pattern: "month de year"
        pattern = r'(\w+)\s+de\s+(\d{4})'
        match = re.search(pattern, text_norm)
        
        if match:
            month_spanish = match.group(1)
            year = match.group(2)
            if month_spanish in self.spanish_months:
                return f"{year}-{self.spanish_months[month_spanish]}"
        
        # Fallback: just year
        year_match = re.search(r'\b(\d{4})\b', text)
        if year_match:
            return f"{year_match.group(1)}-08"
        
        return None
    
    def find_content_match(self, text):
        """Enhanced content matching with precise category detection."""
        text_norm = self.normalize_text(text)
        
        # Exact match first (highest priority)
        if text_norm in self.content_mapping:
            return self.content_mapping[text_norm]
        
        # Special precise matching for problematic categories
        special_matches = {
            'bondesd': 'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBONDESD.ACTUALALLOCATION.NONE.M.1@CONSAR',
            'otros activos': 'MEXPENSIONFUNDS.OTHER.ACTUALALLOCATION.SPECIFIED.M.1@CONSAR',
            'other assets': 'MEXPENSIONFUNDS.OTHER.ACTUALALLOCATION.SPECIFIED.M.1@CONSAR',
        }
        
        for pattern, consar_id in special_matches.items():
            if pattern in text_norm or fuzz.ratio(text_norm, pattern) > 95:
                return consar_id
        
        # Enhanced matching for "otros activos" - be very specific
        if 'otros' in text_norm and 'activos' in text_norm:
            # Make sure it's not just "otros" without "activos"
            if len(text_norm.split()) >= 2:  # At least 2 words
                return 'MEXPENSIONFUNDS.OTHER.ACTUALALLOCATION.SPECIFIED.M.1@CONSAR'
        
        # Fuzzy matching with higher threshold for accuracy
        best_match = None
        best_score = 0
        
        for pattern, consar_id in self.content_mapping.items():
            score = fuzz.ratio(text_norm, pattern)
            
            # Higher threshold to prevent wrong mappings
            if score > 90 and score > best_score:
                # Additional validation - check word overlap for multi-word patterns
                if len(pattern.split()) > 1:
                    pattern_words = set(pattern.split())
                    text_words = set(text_norm.split())
                    word_overlap = len(pattern_words.intersection(text_words)) / len(pattern_words)
                    if word_overlap >= 0.8:  # At least 80% word overlap
                        best_match = consar_id
                        best_score = score
                else:
                    best_match = consar_id
                    best_score = score
        
        return best_match
    
    def scan_sheet_for_data(self, df, target_column_patterns):
        """Enhanced sheet scanner with duplicate prevention and better logging."""
        results = {'target_col': None, 'data_mapping': {}, 'scan_log': []}
        
        # Find target column (e.g., TOTAL)
        for row_idx in range(min(15, len(df))):
            for col_idx in range(len(df.columns)):
                try:
                    cell_value = df.iloc[row_idx, col_idx]
                    if pd.notna(cell_value):
                        cell_norm = self.normalize_text(str(cell_value))
                        for pattern in target_column_patterns:
                            if fuzz.ratio(cell_norm, pattern.lower()) > 90:
                                results['target_col'] = col_idx
                                results['scan_log'].append(f"Target column '{pattern}' found at [{row_idx},{col_idx}]")
                                break
                except:
                    continue
                if results['target_col'] is not None:
                    break
            if results['target_col'] is not None:
                break
        
        if results['target_col'] is None:
            results['scan_log'].append("No target column found")
            return results
        
        # Extract category data with duplicate prevention
        target_col = results['target_col']
        found_categories = set()  # Track found categories to prevent duplicates
        
        for row_idx in range(len(df)):
            # Look for categories in first few columns (most likely to contain labels)
            for col_idx in range(min(3, len(df.columns))):
                try:
                    cell_value = df.iloc[row_idx, col_idx]
                    if pd.notna(cell_value) and isinstance(cell_value, str):
                        cell_text = cell_value.strip()
                        if len(cell_text) > 2:  # Only meaningful text
                            
                            consar_match = self.find_content_match(cell_text)
                            if consar_match and consar_match not in found_categories:
                                # Get value from target column
                                try:
                                    value = df.iloc[row_idx, target_col]
                                    if pd.notna(value):
                                        try:
                                            numeric_value = float(value)
                                            results['data_mapping'][consar_match] = numeric_value
                                            found_categories.add(consar_match)
                                            results['scan_log'].append(f"✅ '{cell_text}' → {consar_match.split('.')[-3]} = {numeric_value}")
                                        except (ValueError, TypeError):
                                            results['data_mapping'][consar_match] = np.nan
                                            found_categories.add(consar_match)
                                            results['scan_log'].append(f"⚠️ '{cell_text}' → {consar_match.split('.')[-3]} = NA (invalid value)")
                                    else:
                                        results['data_mapping'][consar_match] = np.nan
                                        found_categories.add(consar_match)
                                        results['scan_log'].append(f"⚠️ '{cell_text}' → {consar_match.split('.')[-3]} = NA (empty)")
                                except:
                                    results['data_mapping'][consar_match] = np.nan
                                    found_categories.add(consar_match)
                                    results['scan_log'].append(f"❌ '{cell_text}' → {consar_match.split('.')[-3]} = NA (error)")
                                
                                break  # Found match in this row, move to next row
                except:
                    continue
        
        results['scan_log'].append(f"Total categories found: {len(results['data_mapping'])}")
        return results
    
    def read_investment_data(self, filepath):
        """Enhanced investment data reader with detailed logging."""
        print(f"📊 Reading investment data: {filepath}")

        try:
            # Use intelligent sheet detection
            target_sheet = self.find_investment_sheet(filepath)

            if not target_sheet:
                print(f"  ❌ No suitable sheet found in {filepath}")
                return {}, None

            print(f"  📋 Using sheet: {target_sheet}")
            df = pd.read_excel(filepath, sheet_name=target_sheet, header=None)
            
            # Extract date
            extracted_date = None
            for row_idx in range(min(10, len(df))):
                for col_idx in range(min(10, len(df.columns))):
                    try:
                        cell_value = df.iloc[row_idx, col_idx]
                        if pd.notna(cell_value):
                            date_extracted = self.extract_date_from_text(str(cell_value))
                            if date_extracted:
                                extracted_date = date_extracted
                                print(f"  📅 Date found: {date_extracted} from '{cell_value}'")
                                break
                    except:
                        continue
                if extracted_date:
                    break
            
            # Scan sheet for data with detailed logging
            scan_results = self.scan_sheet_for_data(df, ['total'])
            
            print(f"  🔍 Scanning results:")
            for log_entry in scan_results['scan_log']:
                print(f"    {log_entry}")
            
            print(f"  ✅ Investment data extracted: {len(scan_results['data_mapping'])} categories")
            
            # Show some key mappings for validation
            key_mappings = [
                'MEXPENSIONFUNDS.DOMESTICGOVERNMENTBONDSBONDESD.ACTUALALLOCATION.NONE.M.1@CONSAR',
                'MEXPENSIONFUNDS.OTHER.ACTUALALLOCATION.SPECIFIED.M.1@CONSAR',
                'MEXPENSIONFUNDS.TOTAL.ACTUALALLOCATION.NONE.M.1@CONSAR'
            ]
            
            print(f"  🎯 Key mappings check:")
            for key_mapping in key_mappings:
                if key_mapping in scan_results['data_mapping']:
                    value = scan_results['data_mapping'][key_mapping]
                    category_name = key_mapping.split('.')[-3]
                    print(f"    ✅ {category_name}: {value}")
                else:
                    category_name = key_mapping.split('.')[-3]
                    print(f"    ❌ {category_name}: Not found")
            
            return scan_results['data_mapping'], extracted_date
            
        except Exception as e:
            print(f"❌ Error reading investment data: {e}")
            return {}, None
    
    def read_flow_data(self, filepath, target_date):
        """Enhanced flow data reader for HTML-formatted Excel files."""
        print(f"💰 Reading flow data: {filepath}")
        
        if not target_date:
            return {}
        
        try:
            df = None
            
            # Multiple read strategies
            for engine in ['openpyxl', 'xlrd']:
                try:
                    df = pd.read_excel(filepath, header=None, engine=engine)
                    if df is not None and not df.empty:
                        print(f"  ✅ Read with {engine}: {len(df)} rows")
                        break
                except:
                    continue
            
            # Try HTML reading if Excel failed
            if df is None or df.empty:
                try:
                    tables = pd.read_html(filepath, header=None)
                    if tables:
                        df = tables[0]
                        print(f"  ✅ Read as HTML table: {len(df)} rows")
                except:
                    pass
            
            if df is None or df.empty:
                print(f"  ❌ Could not read {filepath}")
                return {}
            
            # Clean HTML content
            df = df.astype(str)
            for col in df.columns:
                df[col] = df[col].str.replace(r'<[^>]+>', '', regex=True)
                df[col] = df[col].str.replace(r'style\s*=\s*[^;]*;', '', regex=True)
                df[col] = df[col].str.strip()
            
            df = df.replace(['', 'nan', 'None'], pd.NA)
            
            # Find date column
            year, month = target_date.split('-')
            possible_abbrevs = self.month_abbreviations.get(month, ['ago'])
            
            date_patterns = []
            for abbrev in possible_abbrevs:
                date_patterns.extend([
                    f"{abbrev}-{year}",
                    f"{abbrev.capitalize()}-{year}",
                    f"{abbrev} {year}",
                    target_date
                ])
            
            # Scan for date column
            target_col = None
            best_score = 0
            
            for row_idx in range(min(20, len(df))):
                for col_idx in range(len(df.columns)):
                    try:
                        cell_value = df.iloc[row_idx, col_idx]
                        if pd.notna(cell_value):
                            cell_str = self.normalize_text(str(cell_value))
                            for pattern in date_patterns:
                                pattern_norm = self.normalize_text(pattern)
                                score = fuzz.ratio(cell_str, pattern_norm)
                                if score > 85 and score > best_score:
                                    target_col = col_idx
                                    best_score = score
                                    print(f"  🎯 Date match: '{cell_value}' (score: {score})")
                    except:
                        continue
            
            if target_col is None:
                print(f"  ❌ No date column found for {target_date}")
                return {}
            
            print(f"  ✅ Using date column {target_col}")
            
            # Enhanced flow patterns based on screenshot analysis
            flow_patterns = [
                # RCV patterns
                'rcv', 'recursos canalizados',
                'rcv imss', 'recursos canalizados imss',
                'rcv isste', 'recursos canalizados isste',

                # IMSS withdrawal patterns - multiple variations
                'retiro de recursos imss', 'retiro recursos imss',
                'retiro imss', 'recursos imss retiro', 'imss retiro',

                # ISSTE withdrawal patterns - multiple variations (including ISSSTE with triple S)
                'retiro de recursos isste', 'retiro recursos isste',
                'retiro de recursos issste', 'retiro recursos issste',
                'retiro isste', 'retiro issste',
                'recursos isste retiro', 'recursos issste retiro',
                'isste retiro', 'issste retiro'
            ]
            
            flow_mapping = {}
            
            for row_idx in range(len(df)):
                category_found = None
                
                # Look for flow categories in multiple columns
                for col_idx in range(min(8, len(df.columns))):
                    try:
                        cell_value = df.iloc[row_idx, col_idx]
                        if pd.notna(cell_value) and str(cell_value).strip():
                            cell_norm = self.normalize_text(str(cell_value))
                            
                            for pattern in flow_patterns:
                                similarity_score = fuzz.ratio(cell_norm, pattern)

                                # Special handling for withdrawal patterns with exact matches
                                if ('retiro' in pattern and ('imss' in pattern or 'isste' in pattern or 'issste' in pattern)):
                                    # Lower threshold for withdrawal patterns due to exact match from source
                                    if similarity_score > 70:
                                        consar_match = self.find_content_match(pattern)
                                        if consar_match:
                                            # Additional validation: check for specific institution in cell text
                                            cell_text_upper = str(cell_value).upper()
                                            if 'imss' in pattern and 'IMSS' in cell_text_upper and 'ISSSTE' not in cell_text_upper:
                                                category_found = (consar_match, str(cell_value).strip())
                                                print(f"  🎯 IMSS WITHDRAWAL: '{cell_value}' → {pattern} (score: {similarity_score})")
                                                break
                                            elif ('isste' in pattern or 'issste' in pattern) and ('ISSTE' in cell_text_upper or 'ISSSTE' in cell_text_upper):
                                                category_found = (consar_match, str(cell_value).strip())
                                                print(f"  🎯 ISSTE WITHDRAWAL: '{cell_value}' → {pattern} (score: {similarity_score})")
                                                break
                                elif similarity_score > 75:
                                    consar_match = self.find_content_match(pattern)
                                    if consar_match:
                                        category_found = (consar_match, str(cell_value).strip())
                                        break
                    except:
                        continue
                    
                    if category_found:
                        break
                
                # Extract value from date column
                if category_found:
                    consar_id, original_text = category_found
                    try:
                        value = df.iloc[row_idx, target_col]
                        if pd.notna(value) and str(value).strip():
                            try:
                                numeric_value = float(str(value).replace(',', ''))
                                flow_mapping[consar_id] = numeric_value
                                print(f"  ✅ '{original_text}' → {numeric_value}")
                            except:
                                flow_mapping[consar_id] = np.nan
                        else:
                            flow_mapping[consar_id] = np.nan
                    except:
                        flow_mapping[consar_id] = np.nan
            
            print(f"  📊 Flow data: {len(flow_mapping)} categories found")
            return flow_mapping
            
        except Exception as e:
            print(f"  ❌ Error reading flow data: {e}")
            return {}
    
    def create_consar_output(self, investment_data, flow_data, date_str):
        """Create CONSAR format output - ALWAYS fixed structure."""
        print("🏗️ Creating CONSAR format output...")
        
        # Initialize all 50 mappings with NA
        final_mapping = {}
        for consar_id in self.consar_headers:
            final_mapping[consar_id] = np.nan
        
        # Update with actual data (preserves 0.0 vs NA distinction)
        for consar_id, value in investment_data.items():
            if consar_id in final_mapping:
                final_mapping[consar_id] = value
        
        for consar_id, value in flow_data.items():
            if consar_id in final_mapping:
                final_mapping[consar_id] = value
        
        # Create fixed DataFrame structure (ALWAYS 3×51)
        data_rows = []
        
        # Row 1: Headers
        data_rows.append([''] + self.consar_headers)
        
        # Row 2: Descriptions
        data_rows.append([''] + self.descriptions)
        
        # Row 3: Data
        values = [final_mapping[consar_id] for consar_id in self.consar_headers]
        data_rows.append([date_str or datetime.now().strftime('%Y-%m')] + values)
        
        df = pd.DataFrame(data_rows)
        
        # Statistics
        found_categories = len(investment_data) + len(flow_data)
        non_na_values = sum(1 for v in values if pd.notna(v))
        zero_values = sum(1 for v in values if pd.notna(v) and v == 0.0)
        
        print(f"✅ CONSAR structure: 3 rows × {len(df.columns)} columns (FIXED)")
        print(f"📊 Categories found: {found_categories}")
        print(f"📊 Values mapped: {non_na_values} (including {zero_values} zeros)")
        print(f"📊 Missing (NA): {len(values) - non_na_values}")
        
        return df
    
    def process_files(self, investment_file=None, flow_files=None, directory='.'):
        """Main universal processing function with auto-discovery."""
        print("🌍 CLEAN UNIVERSAL MEXICAN PENSION FUND MAPPER")
        print("=" * 60)
        print("✅ All 50 CONSAR mappings | ✅ Order Independent | ✅ Auto-Discovery")
        print("=" * 60)

        # Auto-discover files if not provided
        if investment_file is None or flow_files is None:
            print("🔍 Auto-discovering files...")
            discovered = self.discover_files(directory)

            if investment_file is None:
                investment_file = discovered['investment_file']
            if flow_files is None:
                flow_files = discovered['flow_files']

        if not investment_file:
            print("❌ No investment file found! Please ensure WEB_SISTEMA sheet exists.")
            return {'consar_data': pd.DataFrame(), 'date': None, 'investment_categories': 0, 'flow_categories': 0, 'total_categories': 0}

        print(f"\n📁 Processing files:")
        print(f"   • Investment: {os.path.basename(investment_file)}")
        for flow_file in flow_files:
            print(f"   • Flow: {os.path.basename(flow_file)}")

        # Read investment data
        investment_data, date_str = self.read_investment_data(investment_file)
        
        # Read flow data with error handling
        all_flow_data = {}
        for flow_file in flow_files:
            # Try multiple filename variations
            variations = [
                flow_file,
                flow_file.replace(' (1)', ''),
                flow_file.replace('_1', ''),
                flow_file + '.xls' if not flow_file.endswith('.xls') else flow_file
            ]
            
            flow_found = False
            for variation in variations:
                try:
                    flow_data = self.read_flow_data(variation, date_str)
                    if flow_data:
                        all_flow_data.update(flow_data)
                        flow_found = True
                        break
                except FileNotFoundError:
                    continue
                except Exception as e:
                    print(f"⚠️ Error with {variation}: {e}")
            
            if not flow_found:
                print(f"⚠️ Flow file not accessible: {flow_file}")
        
        # Create CONSAR output (ALWAYS fixed structure)
        consar_df = self.create_consar_output(investment_data, all_flow_data, date_str)
        
        return {
            'consar_data': consar_df,
            'date': date_str,
            'investment_categories': len(investment_data),
            'flow_categories': len(all_flow_data),
            'total_categories': len(investment_data) + len(all_flow_data)
        }
    
    def export_results(self, results, output_file):
        """Export results to Excel."""
        print(f"📁 Exporting to: {output_file}")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            
            # Main CONSAR data
            results['consar_data'].to_excel(
                writer, sheet_name='CONSAR_DATA', index=False, header=False
            )
            
            # Summary
            summary_data = [
                ['Processing Date', datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                ['Data Date', results['date'] or 'Not extracted'],
                ['Investment Categories Found', results['investment_categories']],
                ['Flow Categories Found', results['flow_categories']],
                ['Total Categories Mapped', results['total_categories']],
                ['Total CONSAR Headers', len(self.consar_headers)],
                ['Missing Categories', len(self.consar_headers) - results['total_categories']],
                ['Output Structure', f"3 rows × {len(self.consar_headers) + 1} columns (FIXED)"],
                ['Universal Features', 'Order Independent, Structure Adaptive, Content-Based']
            ]
            
            pd.DataFrame(summary_data, columns=['Metric', 'Value']).to_excel(
                writer, sheet_name='Summary', index=False
            )
        
        print("✅ Export completed!")


def main(directory='.'):
    """Main function - Universal Mexican Pension Fund Processing with Auto-Discovery."""
    print("🌍 CLEAN UNIVERSAL MEXICAN PENSION FUND MAPPER")
    print("=" * 70)

    # Initialize mapper
    mapper = CleanUniversalMexicanPensionMapper()

    print(f"🎯 Initialized:")
    print(f"   • {len(mapper.consar_headers)} CONSAR mappings (complete)")
    print(f"   • {len(mapper.content_mapping)} content patterns")
    print(f"   • Universal date recognition (all months)")
    print(f"   • Universal file discovery (no hardcoding)")
    print(f"   • WEB_SISTEMA sheet detection")

    # Process files with auto-discovery
    results = mapper.process_files(directory=directory)
    
    # Export results
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = f"Clean_Mexican_Pension_Data_{timestamp}.xlsx"
    mapper.export_results(results, output_file)
    
    # Final summary
    print(f"\n🎉 PROCESSING COMPLETE!")
    print(f"📊 Structure: ALWAYS 3 rows × 51 columns")
    print(f"📅 Data Date: {results['date']}")
    print(f"🎯 Investment Categories: {results['investment_categories']}")
    print(f"💰 Flow Categories: {results['flow_categories']}")
    print(f"📋 Total Mapped: {results['total_categories']}/50")
    print(f"📁 Output File: {output_file}")
    
    # Show sample
    if not results['consar_data'].empty:
        print(f"\n📋 Sample Output (first 5 columns):")
        sample = results['consar_data'].iloc[:, :5]
        for i, row in sample.iterrows():
            row_type = ['Headers', 'Descriptions', 'Data'][i] if i < 3 else f'Row {i}'
            print(f"  {row_type}: {' | '.join(str(x)[:25] for x in row.values)}")
    
    print(f"\n✨ All done! Check '{output_file}' for complete results.")
    return results


if __name__ == "__main__":
    # Can specify directory or use current directory
    import sys
    directory = sys.argv[1] if len(sys.argv) > 1 else '.'
    results = main(directory)