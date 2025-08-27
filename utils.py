"""
Utility Functions and Classes
Refinery Crude Oil Scheduling System - 12 Tanks Management
FIXED VERSION - Suspension Stock and Time Tracking Corrected
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta, date
import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import tempfile

def get_date_with_ordinal(date_obj):
    """Formats a date object into a string like '17th September'."""
    day = date_obj.day
    if 11 <= day <= 13:
        suffix = 'th'
    else:
        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
    return date_obj.strftime(f'%d{suffix} %B')

def _parse_json_datetime(dt_val):
    """Safely parse a datetime that may have been converted to string via JSON."""
    if not dt_val:
        return None
    if not isinstance(dt_val, str):
        return dt_val
    try:
        # First, try the format used in the cargo report
        return datetime.strptime(dt_val, "%d/%m/%y %H:%M")

    except ValueError:
        try:
            return datetime.strptime(dt_val.replace(' GMT', ''), "%a, %d %b %Y %H:%M:%S")
        except ValueError:
            try:
                return datetime.fromisoformat(dt_val)
            except ValueError:
                 return None

def _save_excel_with_conflict_handling(workbook, base_filename, project_folder=None):
    """Save Excel file with automatic conflict handling"""
    if project_folder is None:
        project_folder = os.getcwd()
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"{base_filename}_{timestamp}.xlsx"
    filepath = os.path.join(project_folder, filename)
    
    counter = 1
    while True:
        try:
            workbook.save(filepath)
            return filename
        except PermissionError:
            filename = f"{base_filename}_{timestamp}_{counter}.xlsx"
            filepath = os.path.join(project_folder, filename)
            counter += 1
            if counter > 20:
                import random
                filename = f"{base_filename}_{timestamp}_{random.randint(1000, 9999)}.xlsx"
                filepath = os.path.join(project_folder, filename)
                workbook.save(filepath)
                return filename
        except Exception as e:
            filename = f"{base_filename}_{timestamp}_alt.xlsx"
            filepath = os.path.join(project_folder, filename)
            workbook.save(filepath)
            return filename

def populate_tank_times(status, tank_id, day_data, feeding_events_log, filling_events_log, tank_object=None):
    """
    Use direct tank datetime values to get start/end times for a given status on a given day,
    and correctly show full-day timings for multi-day events.
    FIXED: Properly display suspended times only on the days they occur
    """
    start_time = ""
    end_time = ""
    
    current_date_str = day_data.get('date', '')
    if not current_date_str or not tank_object:
        return start_time, end_time
    
    def safe_format_time(dt_value):
        return dt_value.strftime('%H:%M') if dt_value else ""
    
    def is_same_date(dt_value, target_date_str):
        return dt_value.strftime('%d/%m/%y') == target_date_str if dt_value else False

    multiday_status_map = {
        'SETTLING': ('settling_start_datetime', 'settling_end_datetime'),
        'LAB_TESTING': ('lab_testing_start_datetime', 'lab_testing_end_datetime')
    }

    if status in multiday_status_map:
        start_attr, end_attr = multiday_status_map[status]
        start_dt = tank_object.get(start_attr)
        end_dt = tank_object.get(end_attr)
        
        if start_dt and end_dt:
            current_day_obj = datetime.strptime(current_date_str, '%d/%m/%y').date()
            
            if start_dt.date() <= current_day_obj <= end_dt.date():
                
                if start_dt.date() == current_day_obj:
                    start_time = safe_format_time(start_dt)
                else:
                    start_time = "00:00"
                
                if end_dt.date() == current_day_obj:
                    end_time = safe_format_time(end_dt)
                else:
                    end_time = "24:00"

                if start_dt.date() == end_dt.date():
                    start_time = safe_format_time(start_dt)
                    end_time = safe_format_time(end_dt)
        
        return start_time, end_time

    # FIXED: Properly handle SUSPENDED status times - only show on actual start/end days
    if status == 'SUSPENDED':
        suspended_start_dt = tank_object.get('suspended_start_datetime')
        suspended_end_dt = tank_object.get('suspended_end_datetime')
        
        if suspended_start_dt:
            # Only show start time on the day suspension actually started
            if is_same_date(suspended_start_dt, current_date_str):
                start_time = safe_format_time(suspended_start_dt)
            
            # Only show end time on the day suspension actually ended
            if suspended_end_dt and is_same_date(suspended_end_dt, current_date_str):
                end_time = safe_format_time(suspended_end_dt)
        
        return start_time, end_time

    status_map = {
        'FEEDING': ('feeding_start_datetime', 'feeding_end_datetime'),
        'EMPTY': ('empty_datetime', None),
        'FILLING': ('filling_start_datetime', 'filling_end_datetime'),
        'FILLED': (None, 'filled_datetime'),
        'READY': ('ready_start_datetime', None)
    }

    if status in status_map:
        start_attr, end_attr = status_map[status]
        if start_attr:
            start_dt = tank_object.get(start_attr)
            if is_same_date(start_dt, current_date_str):
                start_time = safe_format_time(start_dt)
        if end_attr:
            end_dt = tank_object.get(end_attr)
            if is_same_date(end_dt, current_date_str):
                end_time = safe_format_time(end_dt)

    return start_time, end_time

def _calculate_timestamp_consumption_summary(scheduler_instance, params):
    """Calculate timestamp-based consumption summary"""
    summary = {
        'calculation_method': 'Timestamp-based consumption calculation',
        'formula': '(tank_end_time - tank_start_time) / 24 * processing_rate',
        'key_points': [
            'Crude processing date must include timestamp',
            'End feeding time will be 00:00 hrs of next day',
            'Start time for next day will be 00:00 hrs',
            'Consumption calculated as (tank_end_time - tank_start_time) / 24 * processing_rate',
            'SUSPENDED tanks have consumption = 0 with calculated end time'
        ],
        'tank_consumption_details': []
    }
    
    processing_rate = float(params.get('processingRate', 50000))
    
    num_tanks = int(params.get('numTanks', 12))
    for tank_id in range(1, num_tanks + 1):
        tank_consumptions = []
        for day_data in scheduler_instance.simulation_data:
            tank_status = day_data.get(f'tank{tank_id}_status', '')
            start_time = day_data.get(f'tank{tank_id}_status_start_time', '')
            end_time = day_data.get(f'tank{tank_id}_status_end_time', '')
            consumption = day_data.get(f'tank{tank_id}_consumption', 0)
            
            if tank_status in ['FEEDING', 'SUSPENDED'] and (consumption > 0 or tank_status == 'SUSPENDED'):
                tank_consumptions.append({
                    'day': day_data['day'],
                    'status': tank_status,
                    'start_time': start_time,
                    'end_time': end_time,
                    'consumption': consumption,
                    'calculation_note': f"{'SUSPENDED: consumption=0' if tank_status == 'SUSPENDED' else f'({end_time} - {start_time}) / 24 * {processing_rate:,.0f} bbl/day'}"
                })
        
        if tank_consumptions:
            summary['tank_consumption_details'].append({
                'tank_id': tank_id,
                'feeding_days': len([c for c in tank_consumptions if c['status'] == 'FEEDING']),
                'suspended_days': len([c for c in tank_consumptions if c['status'] == 'SUSPENDED']),
                'total_consumption': sum(c['consumption'] for c in tank_consumptions),
                'daily_breakdown': tank_consumptions
            })
    
    return summary

class AdvancedRefineryCrudeScheduler:
    def __init__(self):
        self.simulation_data = []
        self.alerts = []
        self.cargo_schedule = []
        self.emptied_tanks_order = []
        self.initial_params = {}
        self.initial_tank_levels = {}
        self.active_cargos = {}
        self.full_tank_details = []
        self.feeding_events_log = []
        self.filling_events_log = []
        self.daily_discharge_log = []
        
    def _get_processing_start_datetime(self, params):
        """Get processing start datetime, adding a default timestamp if missing."""
        try:
            date_str = params.get('crudeProcessingDate')
            if date_str:
                if 'T' in date_str:
                    return datetime.fromisoformat(date_str.replace('T', ' '))
                else:
                    if ':' not in date_str:
                        date_str += " 08:00"
                        print(f"INFO: No timestamp in crudeProcessingDate. Defaulting to 08:00.")

                    datetime_formats = [
                        "%Y-%m-%d %H:%M",
                        "%d/%m/%y %H:%M",
                        "%d/%m/%Y %H:%M",
                        "%m/%d/%y %H:%M",
                        "%m/%d/%Y %H:%M",
                        "%d-%m-%Y %H:%M",
                        "%d-%m-%y %H:%M",
                    ]
                    for fmt in datetime_formats:
                        try:
                            return datetime.strptime(date_str, fmt)
                        except ValueError:
                            continue
            
            combined = params.get('processingStartDateTime')
            if combined:
                return self._parse_datetime_input(combined)

            date_part = params.get('processingStartDate')
            time_part = params.get('processingStartTime', '08:00')
            if date_part:
                return self._parse_datetime_input(date_part, time_part)
        except Exception:
            pass
        
        default_time = datetime.now().replace(hour=8, minute=0, second=0, microsecond=0)
        print(f"WARNING: Using default processing start time {default_time.strftime('%H:%M')}")
        return default_time

    def _parse_datetime_input(self, date_str, time_str="08:00"):
        """Parse date and time inputs into datetime object"""
        try:
            if 'T' in date_str:
                return datetime.fromisoformat(date_str.replace('T', ' '))
            
            candidate = f"{date_str} {time_str}"
            formats = [
                "%Y-%m-%d %H:%M",
                "%d/%m/%y %H:%M",
                "%d/%m/%Y %H:%M",
                "%m/%d/%y %H:%M",
                "%m/%d/%Y %H:%M",
                "%d-%m-%Y %H:%M",
                "%d-%m-%y %H:%M",
            ]
            for fmt in formats:
                try:
                    return datetime.strptime(candidate, fmt)
                except Exception:
                    continue
        except Exception:
            pass
        return datetime.now().replace(hour=8, minute=0, second=0, microsecond=0)
    
    def _format_datetime_output(self, dt):
        """Format datetime to DD/MM/YY HH:MM"""
        return dt.strftime("%d/%m/%y %H:%M")
    
    def _calculate_buffer_stock(self, params):
        """Calculate buffer stock required for continuous operation"""
        processing_rate = float(params.get('processingRate', 50000))
        pre_journey_days = float(params.get('preJourneyDays', 1))
        journey_days = float(params.get('journeyDays', 10))
        pre_discharge_days = float(params.get('preDischargeDays', 1))
        settling_days = float(params.get('settlingTime', 2))
        lab_testing_days = float(params.get('labTestingDays', 1))
        buffer_days = float(params.get('bufferDays', 2))
        pumping_rate = float(params.get('pumpingRate', 30000))
        
        vlcc_capacity = float(params.get('vlccCapacity', 0))
        suezmax_capacity = float(params.get('suezmaxCapacity', 0))
        aframax_capacity = float(params.get('aframaxCapacity', 0))
        panamax_capacity = float(params.get('panamaxCapacity', 0))
        handymax_capacity = float(params.get('handymaxCapacity', 0))
        largest_cargo = max(vlcc_capacity, suezmax_capacity, aframax_capacity, panamax_capacity, handymax_capacity)
        pumping_days = (largest_cargo / (pumping_rate * 24)) if pumping_rate > 0 and largest_cargo > 0 else 0
        
        lead_time = pre_journey_days + journey_days + pre_discharge_days + pumping_days + settling_days + lab_testing_days
        buffer_stock = (lead_time + buffer_days) * processing_rate
        
        return {
            'lead_time': lead_time,
            'buffer_stock': buffer_stock,
            'components': {
                'pre_journey': pre_journey_days,
                'journey': journey_days,
                'pre_discharge': pre_discharge_days,
                'pumping': pumping_days,
                'settling': settling_days,
                'lab_testing': lab_testing_days,
                'buffer': buffer_days
            }
        }
    
    def _forecast_tank_depletion(self, tanks, processing_rate, current_day):
        """Forecast when tanks will be empty based on current status and processing rate"""
        depletion_forecast = []
        
        for tank in tanks:
            if tank['status'] in ['READY', 'FEEDING'] and tank['available'] > 0:
                days_until_empty = tank['available'] / processing_rate if processing_rate > 0 else 999
                depletion_day = current_day + days_until_empty
                depletion_forecast.append({
                    'tank_id': tank['id'],
                    'depletion_day': depletion_day,
                    'available_volume': tank['available']
                })
        
        depletion_forecast.sort(key=lambda x: x['depletion_day'])
        return depletion_forecast
    
    def _calculate_optimal_arrival_time_for_last_two_tanks(self, tanks, processing_rate, processing_start_dt, current_inventory):
        """
        Calculate when first of last 2 READY tanks will start feeding
        Then work backwards to determine departure date
        """
        if processing_rate <= 0:
            return processing_start_dt + timedelta(days=10)
        
        ready_tanks = [t for t in tanks if t['status'] == 'READY' and t['available'] > 0]
        
        if len(ready_tanks) <= 2:
            return processing_start_dt
        
        tanks_before_last_two = len(ready_tanks) - 2
        total_volume_before_last_two = sum(t['available'] for t in ready_tanks[:tanks_before_last_two])
        
        days_until_last_two = total_volume_before_last_two / processing_rate if processing_rate > 0 else 0
        
        feeding_date_of_first_of_last_two = processing_start_dt + timedelta(days=days_until_last_two)
        
        filled_date_needed = feeding_date_of_first_of_last_two - timedelta(days=3)
        
        return filled_date_needed
    
    def _select_optimal_vessel(self, tanks_empty_count, tank_capacity, available_cargos, processing_rate):
        """
        Selects the optimal vessel based on strategic score.
        Prioritizes vessels that provide more "days of supply".
        """
        if processing_rate <= 0:
            sorted_cargos = sorted(available_cargos.items(), key=lambda x: x[1]['priority'])
            return sorted_cargos[0] if sorted_cargos else None

        best_vessel = None
        best_score = -float('inf')

        total_ullage = tanks_empty_count * tank_capacity

        for cargo_code, cargo_info in available_cargos.items():
            vessel_size = cargo_info['size']
            days_of_supply = vessel_size / processing_rate

            overflow_penalty = 0
            if vessel_size > total_ullage:
                overflow_ratio = (vessel_size - total_ullage) / vessel_size
                overflow_penalty = overflow_ratio * (days_of_supply * 0.25)

            score = days_of_supply - overflow_penalty
            
            if score > best_score:
                best_score = score
                best_vessel = (cargo_code, cargo_info)

        if not best_vessel:
            sorted_cargos = sorted(available_cargos.items(), key=lambda x: x[1]['priority'])
            return sorted_cargos[0] if sorted_cargos else None

        return best_vessel
    
    def _generate_enhanced_cargo_schedule(self, params, departure_mode='solver', lead_time=None):
        """Generate cargo schedule with smart planning for last 2 tanks"""
        schedule = []
        
        processing_rate = float(params.get('processingRate', 50000))
        pumping_rate = float(params.get('pumpingRate', 30000))
        pre_journey_days = float(params.get('preJourneyDays', 1))
        journey_days = float(params.get('journeyDays', 10))
        pre_discharge_days = float(params.get('preDischargeDays', 1))
        settling_days = float(params.get('settlingTime', 2))
        lab_testing_days = float(params.get('labTestingDays', 1))
        report_days = int(params.get('schedulingWindow', 30))
        tank_capacity = float(params.get('tankCapacity', 500000))
        
        # Use provided lead_time or calculate from params
        if lead_time is None:
            lead_time = params.get('leadTime', 15)
        
        available_cargos = {}
        if params.get('vlccCapacity', 0) > 0: 
            available_cargos['vlcc'] = {'size': float(params.get('vlccCapacity')), 'name': 'VLCC', 'priority': 1}
        if params.get('suezmaxCapacity', 0) > 0: 
            available_cargos['suezmax'] = {'size': float(params.get('suezmaxCapacity')), 'name': 'Suezmax', 'priority': 2}
        if params.get('aframaxCapacity', 0) > 0: 
            available_cargos['aframax'] = {'size': float(params.get('aframaxCapacity')), 'name': 'Aframax', 'priority': 3}
        if params.get('panamaxCapacity', 0) > 0: 
            available_cargos['panamax'] = {'size': float(params.get('panamaxCapacity')), 'name': 'Panamax', 'priority': 4}
        if params.get('handymaxCapacity', 0) > 0: 
            available_cargos['handymax'] = {'size': float(params.get('handymaxCapacity')), 'name': 'Handymax', 'priority': 5}
        
        if not available_cargos:
            print("WARNING: No cargo types defined with capacity > 0")
            return []
        
        processing_start_dt = self._get_processing_start_datetime(params)
        
        tanks = []
        total_initial_inventory = 0
        num_tanks = int(params.get('numTanks', 12))
        for i in range(1, num_tanks + 1):
            tank_level = float(params.get(f'tank{i}Level', 0))
        
            dead_bottom = float(params.get(f'deadBottom{i}', 10000))
            available = max(0, tank_level - dead_bottom)
            if available > 0:
                total_initial_inventory += available
            tanks.append({
                'id': i,
                'volume': tank_level,
                'available': available,
                'status': 'READY' if available > 0 else 'EMPTY'
            })
        
        ready_tanks_count = sum(1 for t in tanks if t['status'] == 'READY')
        
        if ready_tanks_count > 2:
            tanks_to_consume = ready_tanks_count - 2
            volume_to_consume = tanks_to_consume * (tank_capacity - 10000)
            days_until_last_two = volume_to_consume / processing_rate if processing_rate > 0 else 10
            
            feeding_date = processing_start_dt + timedelta(days=days_until_last_two)
            
            filled_date = feeding_date - timedelta(days=3)
            
            empty_tanks = sum(1 for t in tanks if t['status'] == 'EMPTY')
            if empty_tanks == 0:
                empty_tanks = 2
                
            vessel_selection = self._select_optimal_vessel(empty_tanks, tank_capacity, available_cargos, processing_rate)
            if vessel_selection:
                cargo_type_code, cargo_info = vessel_selection
                
                pumping_days = cargo_info['size'] / (pumping_rate * 24) if pumping_rate > 0 else 3
                
                pumping_start = filled_date - timedelta(days=pumping_days)
                
                arrival_date = pumping_start - timedelta(days=pre_discharge_days)
                
                departure_date = arrival_date - timedelta(days=journey_days)
                
                dep_back = filled_date
                
                schedule.append({
                    'cargo_id': 1,
                    'type': cargo_info['name'],
                    'size': cargo_info['size'],
                    'dep_port': self._format_datetime_output(departure_date),
                    'arrival': self._format_datetime_output(arrival_date),
                    'dep_back': self._format_datetime_output(dep_back),
                    'pumping_days': round(pumping_days, 1),
                    'departure_datetime': departure_date,
                    'arrival_datetime': arrival_date,
                    'dep_back_datetime': dep_back,
                    'departure_day': (departure_date.date() - processing_start_dt.date()).days + 1,
                    'arrival_day': (arrival_date.date() - processing_start_dt.date()).days + 1
                })
        
        cargo_counter = len(schedule) + 1
        last_arrival = schedule[0]['arrival_datetime'] if schedule else processing_start_dt
        current_inventory = total_initial_inventory
        
        while cargo_counter <= 15:
            vessel_selection = self._select_optimal_vessel(3, tank_capacity, available_cargos, processing_rate)
            if not vessel_selection:
                break
                
                          
            cargo_type_code, cargo_info = vessel_selection
            
            days_of_supply = cargo_info['size'] / processing_rate if processing_rate > 0 else 30
            inventory_ratio = current_inventory / cargo_info['size'] if cargo_info['size'] > 0 else 1
            if inventory_ratio > 2:
                spacing_factor = 1.1
            elif inventory_ratio < 1:
                spacing_factor = 0.4
            else:
                spacing_factor = 0.75
            next_arrival = last_arrival + timedelta(days=days_of_supply * spacing_factor)
            
            pumping_days = cargo_info['size'] / (pumping_rate * 24) if pumping_rate > 0 else 3
            departure_time = next_arrival - timedelta(days= journey_days + pre_discharge_days)
            dep_back_time = next_arrival + timedelta(days=pre_discharge_days + pumping_days)
            
            
            schedule.append({
                'cargo_id': cargo_counter,
                'type': cargo_info['name'],
                'size': cargo_info['size'],
                'dep_port': self._format_datetime_output(departure_time),
                'arrival': self._format_datetime_output(next_arrival),
                'dep_back': self._format_datetime_output(dep_back_time),
                'pumping_days': round(pumping_days, 1),
                'departure_datetime': departure_time,
                'arrival_datetime': next_arrival,
                'dep_back_datetime': dep_back_time,
                'departure_day': (departure_time.date() - processing_start_dt.date()).days + 1,
                'arrival_day': (next_arrival.date() - processing_start_dt.date()).days + 1
            })
            
            last_arrival = next_arrival
            cargo_counter += 1
            current_inventory = max(0, current_inventory - (days_of_supply * spacing_factor * processing_rate))

        return schedule
    
    def _check_cargo_arrival(self, current_day, cargo_schedule):
        """Check if any cargo arrives on the given date"""
        try:
            current_date = current_day.date() if hasattr(current_day, 'date') else current_day
        except Exception:
            current_date = current_day
        for cargo in cargo_schedule:
            arrival_dt = cargo.get('arrival_datetime')
            if arrival_dt and arrival_dt.date() == current_date:
                return {
                    'size': cargo['size'],
                    'type': cargo['type'],
                    'cargo_id': cargo['cargo_id'],
                    'arrival_datetime': cargo.get('arrival_datetime'),
                    'dep_back_datetime': cargo.get('dep_back_datetime')
                }
        return None
    
    def _find_earliest_empty_tank(self, tanks, tanks_feeding_today):
        """Find the earliest emptied tank that is not currently feeding (ALL 12 TANKS)"""
        eligible_tanks = [
            t for t in tanks if
            t['status'] == 'EMPTY' and
            t['id'] not in tanks_feeding_today and
            not t['fed_today']
        ]
        
        if not eligible_tanks:
            return None
        
        eligible_tanks.sort(key=lambda x: x.get('emptied_day', float('inf')))
        return eligible_tanks[0]
    
    def _find_best_feeding_tank(self, tanks, day):
        """Select earliest filled tank (FIFO), not by tank ID order (ALL 12 TANKS)"""
        eligible_tanks = [
            t for t in tanks if
            t['status'] == 'READY' and
            t['available'] > 0 and
            day >= t.get('can_feed_from_day', 1)
        ]
        
        if not eligible_tanks:
            return None
        
        eligible_tanks.sort(key=lambda x: x.get('filled_datetime') or datetime.min)
        
        return eligible_tanks[0]
    
    def _handle_suspended_status(self, tank, current_day, current_date, pumping_rate_per_hour):
        """
        Event-driven SUSPENDED handling:
        - Triggered when filling stops mid-tank due to cargo shortage
        - Set only the SUSPENDED start now; end is set at actual resume
        - Freeze a volume snapshot for correct reporting while suspended
        """
        if tank['status'] != 'FILLING':
            return False
        if (
            tank['volume'] > tank['dead_bottom'] and
            tank['volume'] < tank['capacity'] and
            not tank.get('continuing_fill_tomorrow', False)
        ):
            tank['status'] = 'SUSPENDED'
            # Start = end of the just-finished filling slice (fallback: day start)
            tank['suspended_start_datetime'] = (
                tank.get('filling_end_datetime')
                or datetime.combine(current_date, datetime.min.time())
            )
            # Set end time to 1 hour later (FIX 1)
            tank['suspended_end_datetime'] = tank['suspended_start_datetime'] + timedelta(hours=1)
            # Snapshot the volume at suspension for accurate display while suspended
       
            tank['suspended_volume'] = tank['volume']

            # No consumption while suspended
            tank['daily_consumption'] = 0
            # Announce start only (do not guess an end time)
            self.alerts.append({
                'type': 'warning',
                'day': current_date.strftime('%d/%m'),
                'message': (
                    f'Tank {tank["id"]} SUSPENDED at '
                    f'{tank["suspended_start_datetime"].strftime("%H:%M")} due to cargo shortage.'
                )
            })
            return True
        return False

    def run_simulation(self, params):
        """FIXED: Run simulation with corrected suspension stock tracking and filling from empty display"""
        num_tanks = int(params.get('numTanks', 12))
        
        self.simulation_data = []
        self.alerts = []
        self.cargo_schedule = []
        self.emptied_tanks_order = []
        self.initial_params = params.copy()
        self.initial_tank_levels = {i: float(params.get(f'tank{i}Level', 0)) for i in range(1, num_tanks+1)}
        self.full_tank_details = []
        self.feeding_events_log = []
        self.filling_events_log = []
        self.daily_discharge_log = []
        
        try:
            processing_rate = float(params.get('processingRate', 50000))
            if processing_rate <= 0:
                raise ValueError("Processing rate must be greater than 0")
            
            pumping_rate = float(params.get('pumpingRate', 30000))
            if pumping_rate <= 0:
                raise ValueError("Pumping rate must be greater than 0")
            
            tank_capacity = float(params.get('tankCapacity', 500000))
            if tank_capacity <= 0:
                raise ValueError("Tank capacity must be greater than 0")
            
            processing_rate_per_hour = processing_rate / 24.0
            
            processing_start_dt = self._get_processing_start_datetime(params)
            self.alerts.append({
                'type': 'info', 'day': processing_start_dt.strftime('%d/%m'),
                'message': f'Simulation started on {processing_start_dt.strftime("%d/%m/%y %H:%M")} with processing rate: {processing_rate:,.0f} bbl/day'
            })
            
            departure_mode = params.get('departureMode', 'solver')
            try:
                buffer_info = self._calculate_buffer_stock(params)
                lead_time = buffer_info.get('lead_time', 15)
                self.cargo_schedule = self._generate_enhanced_cargo_schedule(params, departure_mode, lead_time)
            except Exception as e:
                print(f"WARNING: Cargo scheduling failed ({str(e)}), using fallback")
                self.cargo_schedule = []
            
            pumping_rate_per_hour = pumping_rate
            report_days = int(params.get('schedulingWindow', 30))
            disruption_duration = int(params.get('disruptionDuration', 0))
            disruption_start = int(params.get('disruptionStart', 20))
            
            tanks = []
            total_initial_available = 0
            
            for i in range(1, num_tanks + 1):
                tank_level = float(params.get(f'tank{i}Level', 0))
            
                dead_bottom_base = float(params.get(f'deadBottom{i}', 10000))
                buffer_volume = float(params.get('bufferVolume', 500))
                
                dead_bottom_operational = dead_bottom_base + buffer_volume / 2
                
                available_for_inventory = max(0, tank_level - dead_bottom_base)
                available_for_operations = max(0, tank_level - dead_bottom_operational)
                total_initial_available += available_for_inventory
                
                status = 'READY' if tank_level > dead_bottom_operational else 'EMPTY'
                
                tanks.append({
                    'id': i,
                    'volume': tank_level,
                    'status': status,
                    'capacity': tank_capacity,
                    'settling_days_remaining': 0,
                    'dead_bottom': dead_bottom_operational,
                    'dead_bottom_base': dead_bottom_base,
                    'available': available_for_operations,
                    'emptied_day': 0,
                    'daily_consumption': 0,
                    'can_feed_from_day': 1 if status == 'READY' else 0,
                    'fed_today': False,
                    'lab_testing_days_remaining': 0,
                    'feeding_start_datetime': None,
                    'feeding_end_datetime': None,
                    'filling_start_datetime': None,
                    'filling_end_datetime': None,
                    'filled_datetime': None,
                    'settling_start_datetime': None,
                    'settling_end_datetime': None,
                    'lab_testing_start_datetime': None,
                    'lab_testing_end_datetime': None,
                    'ready_start_datetime': None,
                    'empty_datetime': None,
                    'vessel_arrival_datetime': None,
                    'vessel_dep_datetime': None,
                    'filling_cargo_id': None,
                    'original_feeding_start': None,
                    'last_feed_start_volume': 0,
                    'suspended_start_datetime': None,
                    'suspended_end_datetime': None,
                    'suspended_volume': None,
                    'daily_fill_volume': 0,
                    'continuing_fill_tomorrow': False,
                    'emptied_time_today': None,
                    'was_empty_before_filling': False,  # Track if tank was empty before filling
                    'volume_at_day_start': 0,  # Track volume at start of each day
                    'filling_start_volume': 0  # Track exact volume when filling starts
                })
                if status == 'EMPTY':
                    self.emptied_tanks_order.append(i)
            
            active_tank_id = 0
            initial_feed_tank = self._find_best_feeding_tank(tanks, 1)
            if initial_feed_tank:
                active_tank_id = initial_feed_tank['id']
                initial_feed_tank['status'] = 'FEEDING'
                initial_feed_tank['feeding_start_datetime'] = processing_start_dt
                initial_feed_tank['original_feeding_start'] = processing_start_dt
                initial_feed_tank['last_feed_start_volume'] = initial_feed_tank['volume']

                self.alerts.append({
                    'type': 'info', 'day': processing_start_dt.strftime('%d/%m'),
                    'message': f'Initial feeding starts from Tank {active_tank_id} at {processing_start_dt.strftime("%H:%M")}'
                })
                self.feeding_events_log.append({
                    'tank_id': active_tank_id,
                    'start': processing_start_dt,
                    'end': None
                })
            
            base_date = processing_start_dt.replace(hour=0, minute=0, second=0, microsecond=0)
            active_cargo = None
            tanks_emptied_during_day = []
            
            for day in range(1, report_days + 1):
                current_date = (base_date + timedelta(days=day-1)).date()
                actual_date = processing_start_dt + timedelta(days=day-1)
                
                display_day = day
                tanks_emptied_during_day = []
                
                # Track volume at start of day for each tank
                for tank in tanks:
                    tank['volume_at_day_start'] = tank['volume']
                    tank['fed_today'] = False
                    tank['emptied_time_today'] = None
                    tank['daily_consumption'] = 0
                    tank['daily_fill_volume'] = 0
                    tank['continuing_fill_tomorrow'] = False
                    # Reset filling_start_volume at day start if not currently filling
                    if tank['status'] != 'FILLING':
                        tank['filling_start_volume'] = 0
                    
                    if tank['status'] == 'FEEDING' and tank.get('feeding_start_datetime'):
                        if day > 1 and tank.get('original_feeding_start'):
                            if tank['original_feeding_start'].date() < current_date:
                                next_day_start = datetime.combine(current_date, datetime.min.time())
                                tank['feeding_start_datetime'] = next_day_start
                
                tanks_feeding_today = set()
                for tank in tanks:
                    if tank['status'] == 'FEEDING':
                        tanks_feeding_today.add(tank['id'])
                
                hours_elapsed_today = 0.0
                starting_inventory = sum(t['available'] for t in tanks)
                
                day_data = {
                    'day': display_day,
                    'day_index': day,
                    'date': actual_date.strftime('%d/%m/%y'),
                    'arrivals': 0,
                    'cargo_type': '',
                    'processing': 0,
                    'clash_detected': False,
                    'active_tank_id': active_tank_id,
                    'start_inventory': starting_inventory,
                    'cargo_opening_stock': 0,
                    'cargo_consumption_today': 0,
                    'cargo_closing_stock': 0
                }
                
                for tank in tanks:
                    end_of_today = base_date + timedelta(days=day)

                    # FIX 2: SUSPENDED tanks transition to EMPTY after 1 hour
                    if tank['status'] == 'SUSPENDED':
                        if tank.get('suspended_start_datetime'):
                            time_since_suspension = (end_of_today - tank['suspended_start_datetime']).total_seconds() / 3600
                            if time_since_suspension > 1:
                                tank['status'] = 'EMPTY'
                                self.alerts.append({
                                    'type': 'info', 'day': actual_date.strftime('%d/%m'),
                                    'message': f'Tank {tank["id"]} transitioned from SUSPENDED to EMPTY'
                                })

                    if tank['status'] == 'SETTLING':
                        settling_end_dt = tank.get('settling_end_datetime')
                        if settling_end_dt and end_of_today > settling_end_dt:
                            tank['status'] = 'LAB_TESTING'
                            tank['lab_testing_start_datetime'] = settling_end_dt
                            tank['lab_testing_end_datetime'] = settling_end_dt + timedelta(hours=24)
                            tank['daily_consumption'] = 0
                            
                            self.alerts.append({
                                'type': 'info', 'day': actual_date.strftime('%d/%m'),
                                'message': f'Tank {tank["id"]} SETTLING complete at {settling_end_dt.strftime("%H:%M")}, starts LAB_TESTING for 24 hours until {tank["lab_testing_end_datetime"].strftime("%d/%m %H:%M")}'
                            })
                            
                    elif tank['status'] == 'LAB_TESTING':
                        lab_end_dt = tank.get('lab_testing_end_datetime')
                        if lab_end_dt and end_of_today > lab_end_dt:
                            tank['status'] = 'READY'
                            tank['ready_start_datetime'] = lab_end_dt
                            tank['can_feed_from_day'] = day + 1
                            tank['available'] = max(0, tank['volume'] - tank['dead_bottom'])
                            tank['daily_consumption'] = 0

                            for event in reversed(self.filling_events_log):
                                if event['tank_id'] == tank['id'] and event['end'] is None:
                                    event['end'] = tank.get('filling_end_datetime')
                                    event['settle_start'] = tank.get('settling_start_datetime')
                                    event['lab_start'] = tank.get('lab_testing_start_datetime')
                                    event['ready_time'] = lab_end_dt
                                    break
                            
                            available_date = current_date + timedelta(days=1)
                            available_date_str = get_date_with_ordinal(available_date)
                            self.alerts.append({
                                'type': 'success', 'day': actual_date.strftime('%d/%m'),
                                'message': f'Tank {tank["id"]} LAB_TESTING complete at {lab_end_dt.strftime("%H:%M")}, now READY. Available for feeding from {available_date_str}'
                            })
                
                processing_demand_today = processing_rate
                tanks_used_today = set()
                day_start_time = datetime.combine(current_date, datetime.min.time())
                day_end_time = day_start_time + timedelta(days=1)
                
                actual_start_time = processing_start_dt
                if day == 1:
                    if actual_start_time.hour > 0 or actual_start_time.minute > 0:
                        hours_available = 24 - (actual_start_time.hour + actual_start_time.minute/60)
                        processing_demand_today = (hours_available / 24) * processing_rate
                
                # FEEDING LOGIC
                while processing_demand_today > 0:
                    active_tank = next((t for t in tanks if t['id'] == active_tank_id), None)
                    
                    if not active_tank or active_tank['status'] != 'FEEDING':
                        old_tank_id = active_tank_id
                        
                        next_feed_tank = self._find_best_feeding_tank(tanks, day)
                        
                        if next_feed_tank:
                            active_tank_id = next_feed_tank['id']
                            active_tank = next_feed_tank
                            active_tank['status'] = 'FEEDING'
                            
                            start_feed_time = datetime.combine(current_date, datetime.min.time()) + timedelta(hours=hours_elapsed_today)
                            if day == 1 and hours_elapsed_today == 0:
                                start_feed_time = actual_start_time
                                active_tank['original_feeding_start'] = actual_start_time
                            
                            active_tank['feeding_start_datetime'] = start_feed_time
                            active_tank['last_feed_start_volume'] = active_tank['volume']
                            
                            tanks_feeding_today.add(active_tank_id)
                            
                            self.feeding_events_log.append({
                                'tank_id': active_tank_id,
                                'start': start_feed_time,
                                'end': None
                            })
                            
                            if old_tank_id != 0:
                                self.alerts.append({
                                    'type': 'info', 'day': actual_date.strftime('%d/%m'),
                                    'message': f'Switched to feed from Tank {active_tank_id} at {start_feed_time.strftime("%H:%M")}.'
                                })
                        else:
                            active_tank_id = 0
                            break
                    
                    if not active_tank:
                        break
                    
                    tanks_feeding_today.add(active_tank['id'])
                    tanks_used_today.add(active_tank['id'])
                    
                    consumable_volume = max(0, active_tank['volume'] - active_tank['dead_bottom'])
                    amount_to_take = min(processing_demand_today, consumable_volume)
                    
                    if amount_to_take > 0:
                        active_tank['fed_today'] = True
                        active_tank['volume'] -= amount_to_take
                        active_tank['available'] = max(0, active_tank['volume'] - active_tank['dead_bottom'])
                        active_tank['daily_consumption'] += amount_to_take
                        processing_demand_today -= amount_to_take
                        hours_for_this = (amount_to_take / processing_rate_per_hour) if processing_rate_per_hour > 0 else 0
                        hours_elapsed_today += hours_for_this
                    
                    if active_tank['volume'] <= active_tank['dead_bottom']:
                        active_tank['volume'] = active_tank['dead_bottom']
                        active_tank['available'] = 0
                        active_tank['status'] = 'EMPTY'
                        active_tank['emptied_day'] = day
                        
                        end_time = datetime.combine(current_date, datetime.min.time()) + timedelta(hours=hours_elapsed_today)
                        if day == 1 and actual_start_time.hour > 0:
                            end_time = actual_start_time + timedelta(hours=hours_elapsed_today)

                        active_tank['feeding_end_datetime'] = end_time
                        active_tank['empty_datetime'] = end_time
                        active_tank['emptied_time_today'] = end_time
                        tanks_emptied_during_day.append({'tank_id': active_tank['id'], 'time': end_time})
                        
                        if active_tank['id'] not in self.emptied_tanks_order:
                            self.emptied_tanks_order.append(active_tank['id'])
                        
                        consumption = active_tank['last_feed_start_volume'] - active_tank['volume']
                        
                        for event in reversed(self.feeding_events_log):
                            if event['tank_id'] == active_tank['id'] and event['end'] is None:
                                event['end'] = end_time
                                event['start_level'] = active_tank['last_feed_start_volume']
                                event['end_level'] = active_tank['volume']
                                event['consumption'] = consumption
                                break
                        
                        self.alerts.append({
                            'type': 'warning', 'day': actual_date.strftime('%d/%m'),
                            'message': f'Tank {active_tank["id"]} emptied at {end_time.strftime("%H:%M")} (consumed {consumption:,.0f} bbl, {active_tank["volume"]:,.0f} bbl remaining at dead bottom)'
                        })
                        
                        if processing_demand_today > 0:
                            next_tank = self._find_best_feeding_tank(tanks, day)
                            if next_tank:
                                old_tank_id = active_tank['id']
                                active_tank_id = next_tank['id']
                                next_tank['status'] = 'FEEDING'
                                start_time = datetime.combine(current_date, datetime.min.time()) + timedelta(hours=hours_elapsed_today)
                                next_tank['feeding_start_datetime'] = start_time
                                next_tank['last_feed_start_volume'] = next_tank['volume']
                                
                                tanks_feeding_today.add(active_tank_id)

                                self.feeding_events_log.append({
                                    'tank_id': active_tank_id,
                                    'start': start_time,
                                    'end': None
                                })
                                
                                self.alerts.append({
                                    'type': 'info', 'day': actual_date.strftime('%d/%m'),
                                    'message': f'Tank {old_tank_id} emptied. Switched to feed from Tank {active_tank_id} at {start_time.strftime("%H:%M")}.'
                                })
                            else:
                                active_tank_id = 0
                                break
                        else:
                            active_tank_id = 0
                            break
                
                daily_tank_depletion = 0
                for tank in tanks:
                    if tank['daily_consumption'] > 0:
                        daily_tank_depletion += tank['daily_consumption']
                
                if len(tanks_used_today) > 1:
                    tank_consumptions = []
                    total_consumption = 0
                    for tank_id in tanks_used_today:
                        tank = next((t for t in tanks if t['id'] == tank_id), None)
                        if tank and tank['daily_consumption'] > 0:
                            tank_consumptions.append(f"Tank {tank_id}: {tank['daily_consumption']:,.0f}")
                            total_consumption += tank['daily_consumption']
                    
                    self.alerts.append({
                        'type': 'info', 'day': actual_date.strftime('%d/%m'),
                        'message': f'Multiple tanks feeding: {", ".join(tank_consumptions)}. Total: {total_consumption:,.0f} bbl'
                    })
                
                day_data['processing'] = daily_tank_depletion
                day_data['daily_tank_depletion'] = daily_tank_depletion
                day_data['active_tank_id'] = active_tank_id
                
                # CARGO ARRIVAL AND FILLING LOGIC
                if not (disruption_duration > 0 and disruption_start <= day < disruption_start + disruption_duration):
                    arrival_info = self._check_cargo_arrival(current_date, self.cargo_schedule)
                    if arrival_info:
                        if not active_cargo:
                            active_cargo = arrival_info
                            active_cargo['remaining_volume'] = active_cargo['size']
                            active_cargo['pumping_start_time'] = active_cargo['arrival_datetime'] + timedelta(days=float(self.initial_params.get('preDischargeDays', 1)))
                            day_data.update({'arrivals': active_cargo['size'], 'cargo_type': active_cargo['type']})

                            if active_cargo.get('arrival_datetime'):
                                day_data['arrival_datetime_str'] = self._format_datetime_output(active_cargo['arrival_datetime'])

                            self.alerts.append({
                                'type': 'success', 'day': actual_date.strftime('%d/%m'),
                                'message': f"Vessel {active_cargo['type']} arrived. Pumping will begin at {active_cargo['pumping_start_time'].strftime('%d/%m %H:%M')}"
                            })
                        else:
                            if active_cargo['remaining_volume'] <= 0:
                                actual_pumping_end_time = current_pumping_time

                                for schedule_entry in self.cargo_schedule:
                                    if schedule_entry.get('cargo_id') == active_cargo.get('cargo_id'):
                                        schedule_entry['actual_pumping_end_dt'] = actual_pumping_end_time
                                        break
                                
                                
                                active_cargo = arrival_info
                                active_cargo['remaining_volume'] = active_cargo['size']
                                active_cargo['pumping_start_time'] = active_cargo['arrival_datetime'] + timedelta(days=float(self.initial_params.get('preDischargeDays', 1)))
                                day_data.update({'arrivals': active_cargo['size'], 'cargo_type': active_cargo['type']})
                                self.alerts.append({
                                    'type': 'success', 'day': actual_date.strftime('%d/%m'),
                                    'message': f"Vessel {active_cargo['type']} arrived. Previous cargo completed. Pumping will begin at {active_cargo['pumping_start_time'].strftime('%d/%m %H:%M')}"
                                })
                            else:
                                self.alerts.append({
                                    'type': 'warning', 'day': actual_date.strftime('%d/%m'),
                                    'message': f"Vessel {arrival_info['type']} arrived but {active_cargo['type']} still discharging. Will queue next."
                                })
                
                if active_cargo and active_cargo['remaining_volume'] > 0 and current_date >= active_cargo['pumping_start_time'].date():
                    cargo_opening_stock = active_cargo['remaining_volume']
                    cargo_consumption_today = 0
                    
                    pumping_start_this_day = max(datetime.combine(current_date, datetime.min.time()), active_cargo['pumping_start_time'])
                    hours_to_pump_today = (day_end_time - pumping_start_this_day).total_seconds() / 3600
                    volume_to_pump_today = min(hours_to_pump_today * pumping_rate_per_hour, active_cargo['remaining_volume'])
                    current_pumping_time = pumping_start_this_day
                    
                    tanks_available_for_filling = False
                    
                    while volume_to_pump_today > 0 and active_cargo['remaining_volume'] > 0:
                        currently_filling = sum(1 for t in tanks if t['status'] == 'FILLING')
                        
                        if currently_filling >= 2:
                            self.alerts.append({
                                'type': 'warning', 'day': actual_date.strftime('%d/%m'),
                                'message': f"Filling limit reached (2 tanks). Waiting for tank to complete filling."
                            })
                            break
                        
                        # FIX 3: Skip suspended tanks - they need to transition to EMPTY first
                        target_tank = None
                        
                        # Look for already FILLING tanks first
                        target_tank = next((t for t in tanks if t['status'] == 'FILLING'), None)
                        
                        if not target_tank:
                            for emptied_info in tanks_emptied_during_day:
                                if pumping_start_this_day >= emptied_info['time']:
                                    potential_tank = next((t for t in tanks if t['id'] == emptied_info['tank_id'] and t['status'] == 'EMPTY'), None)
                                    if potential_tank and potential_tank['id'] not in tanks_feeding_today and not potential_tank['fed_today']:
                                        target_tank = potential_tank
                                        # Store exact volume at moment of starting to fill
                                        target_tank['filling_start_volume'] = target_tank['volume']  # Should be dead_bottom
                                        target_tank['was_empty_before_filling'] = True
                                        tanks_available_for_filling = True
                                        break
                            
                            if not target_tank:
                                target_tank = self._find_earliest_empty_tank(tanks, tanks_feeding_today)
                                if target_tank:
                                    # Store exact volume at moment of starting to fill
                                    target_tank['filling_start_volume'] = target_tank['volume']  # Should be dead_bottom
                                    target_tank['was_empty_before_filling'] = True
                                    tanks_available_for_filling = True
                        else:
                            tanks_available_for_filling = True
                        
                        if target_tank:
                            if target_tank['status'] != 'FILLING':
                                start_fill_time = current_pumping_time
                                # Store exact volume at the moment filling starts
                                if not target_tank.get('filling_start_volume'):
                                    target_tank['filling_start_volume'] = target_tank['volume']
                                target_tank['filling_start_datetime'] = start_fill_time
                                target_tank['filling_cargo_id'] = active_cargo['cargo_id']
                                target_tank['vessel_arrival_datetime'] = active_cargo.get('arrival_datetime')
                                target_tank['vessel_dep_datetime'] = active_cargo.get('dep_back_datetime')
                                
                                self.filling_events_log.append({
                                    'tank_id': target_tank['id'],
                                    'start': start_fill_time,
                                    'end': None,
                                    'settle_start': None,
                                    'lab_start': None,
                                    'ready_time': None
                                })

                                self.alerts.append({
                                    'type': 'info', 'day': actual_date.strftime('%d/%m'),
                                    'message': f"Started filling Tank {target_tank['id']} at {start_fill_time.strftime('%H:%M')}"
                                })
                        else:
                            cargo_departed = False
                            if active_cargo.get('dep_back_datetime'):
                                if active_cargo['dep_back_datetime'].date() <= current_date:
                                    cargo_departed = True
                            
                            if not cargo_departed and not tanks_available_for_filling:
                                self.alerts.append({
                                    'type': 'danger', 'day': actual_date.strftime('%d/%m'),
                                    'message': f"DEMURRAGE: {active_cargo['type']} waiting - no empty tank available"
                                })
                            break
                        
                        space_in_tank = tank_capacity - target_tank['volume']
                        volume_for_this_tank = min(space_in_tank, volume_to_pump_today, active_cargo['remaining_volume'])
                        
                        if volume_for_this_tank > 0:
                            self.daily_discharge_log.append({
                                'date': actual_date.strftime('%d/%m/%y'),
                                'cargo_type': active_cargo['type'],
                                'tank_id': target_tank['id'],
                                'volume_filled': volume_for_this_tank
                            })
                            
                            target_tank['status'] = 'FILLING'
                            target_tank['volume'] += volume_for_this_tank
                            target_tank['daily_fill_volume'] += volume_for_this_tank
                            target_tank['daily_consumption'] = -(target_tank['daily_fill_volume'])
                            active_cargo['remaining_volume'] -= volume_for_this_tank
                            volume_to_pump_today -= volume_for_this_tank
                            cargo_consumption_today += volume_for_this_tank
                            
                            pumping_hours = volume_for_this_tank / pumping_rate_per_hour if pumping_rate_per_hour > 0 else 0
                            current_pumping_time += timedelta(hours=pumping_hours)
                            
                            if target_tank['volume'] >= tank_capacity - 1:
                                target_tank['volume'] = tank_capacity
                                filling_end_time = current_pumping_time
                                target_tank['filling_end_datetime'] = filling_end_time
                                
                                target_tank['status'] = 'FILLED'
                                target_tank['filled_datetime'] = filling_end_time
                                target_tank['daily_consumption'] = 0
                                
                                target_tank['status'] = 'SETTLING'
                                target_tank['settling_start_datetime'] = filling_end_time
                                target_tank['settling_end_datetime'] = filling_end_time + timedelta(hours=48)
                                
                                self.alerts.append({
                                    'type': 'info', 'day': actual_date.strftime('%d/%m'),
                                    'message': f"Tank {target_tank['id']} FILLED at {filling_end_time.strftime('%H:%M')}, starts SETTLING for 48 hours until {target_tank['settling_end_datetime'].strftime('%d/%m %H:%M')}"
                                })
                                
                                if volume_to_pump_today > 0 and active_cargo['remaining_volume'] > 0:
                                    next_tank = self._find_earliest_empty_tank(tanks, tanks_feeding_today)
                                    if next_tank:
                                        next_tank['status'] = 'FILLING'
                                        # Store exact volume at the moment filling starts
                                        next_tank['filling_start_volume'] = next_tank['volume']
                                        next_tank['was_empty_before_filling'] = True
                                        start_fill_time = current_pumping_time
                                        next_tank['filling_start_datetime'] = start_fill_time
                                        next_tank['filling_cargo_id'] = active_cargo['cargo_id']
                                        next_tank['vessel_arrival_datetime'] = active_cargo.get('arrival_datetime')
                                        next_tank['vessel_dep_datetime'] = active_cargo.get('dep_back_datetime')
                                        
                                        self.filling_events_log.append({
                                            'tank_id': next_tank['id'],
                                            'start': start_fill_time,
                                            'end': None,
                                            'settle_start': None,
                                            'lab_start': None,
                                            'ready_time': None
                                        })
                                        
                                        self.alerts.append({
                                            'type': 'info', 'day': actual_date.strftime('%d/%m'),
                                            'message': f"Seamless handoff: Started filling Tank {next_tank['id']} at {current_pumping_time.strftime('%H:%M')}"
                                        })
                        
                        if active_cargo['remaining_volume'] <= 0:  
                            actual_pumping_end_time = current_pumping_time 
                            filling_end_time = current_pumping_time
                            target_tank['filling_end_datetime'] = filling_end_time
                            
                            if (target_tank['volume'] > target_tank['dead_bottom'] and 
                                target_tank['volume'] < tank_capacity):
                                
                                # FIXED: Calculate suspended volume based on actual filling time from filling_start_datetime
                                filling_start_dt = target_tank.get('filling_start_datetime')
                                if filling_start_dt:
                                    # Calculate actual hours pumped from filling start to suspension
                                    hours_pumped = (filling_end_time - filling_start_dt).total_seconds() / 3600
                                    volume_pumped = hours_pumped * pumping_rate_per_hour
                                    
                                    # Suspended volume = volume when filling started + what was actually pumped
                                    suspended_volume = target_tank.get('filling_start_volume', target_tank['dead_bottom']) + volume_pumped
                                    target_tank['suspended_volume'] = suspended_volume
                                else:
                                    # Fallback: use current volume
                                    target_tank['suspended_volume'] = target_tank['volume']
                                
                                target_tank['status'] = 'SUSPENDED'
                                target_tank['suspended_start_datetime'] = filling_end_time
                                # FIX 4: Set end time to 1 hour later
                                target_tank['suspended_end_datetime'] = filling_end_time + timedelta(hours=1)
                                target_tank['daily_consumption'] = 0
                                
                                self.alerts.append({
                                    'type': 'warning', 'day': actual_date.strftime('%d/%m'),
                                    'message': f"Tank {target_tank['id']} SUSPENDED at {filling_end_time.strftime('%H:%M')} due to cargo shortage. Volume at suspension: {target_tank['suspended_volume']:,.0f} bbl. Will transition to EMPTY after 1 hour."
                                })
                    
                    cargo_closing_stock = active_cargo['remaining_volume']
                    day_data.update({
                        'cargo_opening_stock': cargo_opening_stock,
                        'cargo_consumption_today': cargo_consumption_today,
                        'cargo_closing_stock': cargo_closing_stock
                    })
                    
                    if active_cargo['remaining_volume'] <= 0:
                        actual_pumping_end_time = current_pumping_time

                        self.alerts.append({
                            'type': 'success', 'day': actual_date.strftime('%d/%m'),
                            'message': f"Cargo {active_cargo['type']} completely discharged"
                        })
                        active_cargo = None
                
                ending_inventory = sum(t['available'] for t in tanks)
                day_data['end_inventory'] = ending_inventory
                total_usable_capacity = sum(tank_capacity for t in tanks)
                day_data['tank_utilization'] = (ending_inventory / total_usable_capacity) * 100 if total_usable_capacity > 0 else 0
                
                # FIXED: Stock calculation for suspended tanks and filling from empty
                for tank in tanks:
                    # Calculate opening and closing stocks properly
                    if tank['status'] == 'SUSPENDED' and tank.get('suspended_volume') is not None:
                        # For suspended tanks, closing stock is the suspended volume
                        closing_stock = tank['suspended_volume']
                        # Opening stock calculation
                        if tank.get('suspended_start_datetime') and tank['suspended_start_datetime'].date() == current_date:
                            # Tank got suspended today, opening stock is volume when filling started
                            opening_stock = tank.get('filling_start_volume', tank['volume_at_day_start'])
                        else:
                            # Tank was already suspended from previous day
                            opening_stock = tank.get('suspended_volume', tank['volume_at_day_start'])
                    elif tank['status'] == 'FILLING':
                        # Check if this tank started filling today
                        if tank.get('filling_start_datetime') and tank['filling_start_datetime'].date() == current_date:
                            # Tank started filling today
                            if tank.get('was_empty_before_filling'):
                                # FIXED: If tank was empty before filling, show dead bottom for both opening and closing
                                opening_stock = tank.get('filling_start_volume', tank['dead_bottom'])
                                closing_stock = tank['volume']
                            else:
                                # Tank was partially filled before
                                opening_stock = tank.get('filling_start_volume', tank['volume'] - tank['daily_fill_volume'])
                                closing_stock = tank['volume']
                        else:
                            # Tank was already filling from previous day
                            opening_stock = tank['volume_at_day_start']
                            closing_stock = tank['volume']
                    else:
                        # Normal calculation for other statuses
                        opening_stock = tank['volume'] + tank['daily_consumption'] - tank['daily_fill_volume']
                        closing_stock = tank['volume']
                    
                    day_data.update({
                        f'tank{tank["id"]}_level': tank['volume'],
                        f'tank{tank["id"]}_status': tank['status'],
                        f'tank{tank["id"]}_consumption': tank['daily_consumption'],
                        f'tank{tank["id"]}_opening_stock': opening_stock,
                        f'tank{tank["id"]}_closing_stock': closing_stock,
                        f'tank{tank["id"]}_status_start_time': '',
                        f'tank{tank["id"]}_status_end_time': '',
                        f'tank{tank["id"]}_filling_cargo': tank.get('filling_cargo_id', ''),
                        f'tank{tank["id"]}_filled_time': '',
                        f'tank{tank["id"]}_suspended_start': '',
                        f'tank{tank["id"]}_suspended_end': ''
                    })
                    
                    start_time, end_time = populate_tank_times(tank['status'], tank['id'], day_data, self.feeding_events_log, self.filling_events_log, tank)
                    day_data[f'tank{tank["id"]}_status_start_time'] = start_time
                    day_data[f'tank{tank["id"]}_status_end_time'] = end_time
                    
                    if tank['status'] == 'SUSPENDED':
                        if tank.get('suspended_start_datetime') and tank['suspended_start_datetime'].date() == current_date:
                            day_data[f'tank{tank["id"]}_suspended_start'] = tank['suspended_start_datetime'].strftime('%H:%M')
                        if tank.get('suspended_end_datetime') and tank['suspended_end_datetime'].date() == current_date:
                            day_data[f'tank{tank["id"]}_suspended_end'] = tank['suspended_end_datetime'].strftime('%H:%M')
                    
                    if tank['status'] == 'SETTLING':
                        if tank.get('settling_start_datetime') and tank['settling_start_datetime'].date() == current_date:
                            day_data[f'tank{tank["id"]}_filled_time'] = tank['settling_start_datetime'].strftime('%H:%M')
                
                self.simulation_data.append(day_data)
            
            self.full_tank_details = tanks
            
            metrics = self._calculate_metrics(params)
            buffer_info = self._calculate_buffer_stock(params)
            cargo_report = self._generate_cargo_report(params, self.cargo_schedule)
            
            final_feeding_end_dt = None
            for tank in tanks:
                if tank.get('feeding_end_datetime'):
                    if final_feeding_end_dt is None or tank['feeding_end_datetime'] > final_feeding_end_dt:
                        final_feeding_end_dt = tank['feeding_end_datetime']
            
            first_feeding_start_dt = None
            for tank in tanks:
                if tank.get('original_feeding_start'):
                    if first_feeding_start_dt is None or tank['original_feeding_start'] < first_feeding_start_dt:
                        first_feeding_start_dt = tank['original_feeding_start']
            
            first_filling_start_dt = None
            last_filling_end_dt = None
            for tank in tanks:
                if tank.get('filling_start_datetime'):
                    if first_filling_start_dt is None or tank['filling_start_datetime'] < first_filling_start_dt:
                        first_filling_start_dt = tank['filling_start_datetime']
                if tank.get('filling_end_datetime'):
                    if last_filling_end_dt is None or tank['filling_end_datetime'] > last_filling_end_dt:
                        last_filling_end_dt = tank['filling_end_datetime']
            
            initial_start_time_str = self._format_datetime_output(first_feeding_start_dt) if first_feeding_start_dt else "N/A"
            final_end_time_str = self._format_datetime_output(final_feeding_end_dt) if final_feeding_end_dt else "N/A"
            first_filling_start_str = self._format_datetime_output(first_filling_start_dt) if first_filling_start_dt else "N/A"
            last_filling_end_str = self._format_datetime_output(last_filling_end_dt) if last_filling_end_dt else "N/A"
            
           

            return {
                'parameters': params,
                'simulation_data': self.simulation_data,
                'alerts': self.alerts,
                'metrics': metrics,
                'cargo_schedule': self.cargo_schedule,
                'cargo_report': cargo_report,
                'feeding_events_log': self.feeding_events_log,
                'filling_events_log': self.filling_events_log,
                'daily_discharge_log': self.daily_discharge_log,
                'buffer_info': buffer_info,
                'initial_start_time': initial_start_time_str,
                'final_end_time': final_end_time_str,
                'first_filling_start_time': first_filling_start_str,
                'last_filling_end_time': last_filling_end_str,
                'full_tank_details': self.full_tank_details,
            }
            
        except ZeroDivisionError as e:
            return {'error': f'Division by zero error: {str(e)}. Please check input parameters'}
        except Exception as e:
            import traceback
            traceback.print_exc()
            return {'error': str(e)}
    
    def _calculate_metrics(self, params):
        # [UNCHANGED - keeping exact same implementation]
        """Calculate simulation metrics"""
        if not self.simulation_data:
            return {}
        
        processing_rate = float(params.get('processingRate', 50000))
        total_processed = sum(day['processing'] for day in self.simulation_data)
        inventories = [day['end_inventory'] for day in self.simulation_data]
        
        return {
            'total_processed': total_processed,
            'avg_utilization': np.mean([day['tank_utilization'] for day in self.simulation_data]) if self.simulation_data else 0,
            'min_inventory': min(inventories) if inventories else 0,
            'max_inventory': max(inventories) if inventories else 0,
            'critical_days': sum(1 for day in self.simulation_data if day['end_inventory'] < processing_rate * 3),
            'clash_days': sum(1 for day in self.simulation_data if day.get('clash_detected', False)),
            'processing_efficiency': (total_processed / (processing_rate * len(self.simulation_data))) * 100 if processing_rate > 0 and self.simulation_data else 0,
            'sustainable_processing': min(inventories) >= 0 if inventories else False,
            'avg_processing_rate': total_processed / len(self.simulation_data) if self.simulation_data else 0,
            'inventory_trend': inventories,
            'total_cargoes': len([day for day in self.simulation_data if day['arrivals'] > 0]),
            'cargo_mix': ', '.join([
                f"{len([d for d in self.simulation_data if d.get('cargo_type') == ct])} {ct}"
                for ct in set(d.get('cargo_type', '') for d in self.simulation_data if d.get('cargo_type'))
            ])
        }
    
    def _generate_cargo_report(self, params, cargo_schedule):
        # [UNCHANGED - keeping exact same implementation]
        """Generate detailed cargo report using the original, accurate cargo_schedule"""
        cargo_report = []
        if not cargo_schedule:
            return cargo_report

        pre_journey_days = float(params.get('preJourneyDays', 1))
        processing_start_dt = self._get_processing_start_datetime(params)

        for cargo in cargo_schedule:
            try:
                arrival_dt = cargo.get('arrival_datetime')
                departure_dt = cargo.get('departure_datetime')
                
                actual_pumping_end_dt = cargo.get('actual_pumping_end_dt')
                planned_dep_unload_port_dt = cargo.get('dep_back_datetime')
                dep_unload_port_dt = actual_pumping_end_dt if actual_pumping_end_dt else planned_dep_unload_port_dt

                load_port_time_dt = departure_dt - timedelta(days=pre_journey_days) if departure_dt else None
                arrival_day = (arrival_dt.date() - processing_start_dt.date()).days + 1 if arrival_dt else 'N/A'
                cargo_size = cargo.get('size', 0)

                cargo_report.append({
                    'Day': arrival_day,
                    'Cargo_type': cargo.get('type', ''),
                    'Load_Port_time': self._format_datetime_output(load_port_time_dt) if load_port_time_dt else '',
                    'dep_time': self._format_datetime_output(departure_dt) if departure_dt else '',
                    'Arrival_time': self._format_datetime_output(arrival_dt) if arrival_dt else '',
                    'dep_unload_port': self._format_datetime_output(dep_unload_port_dt) if dep_unload_port_dt else '',
                    'Cargo_size': f"{cargo_size:,.0f} bbl"
                })
            except Exception as e:
                print(f"Error processing cargo for report: {e}")
                continue
        
        return cargo_report

from flask import Flask

app = Flask(__name__)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)