"""
Utility Functions and Classes
Refinery Crude Oil Scheduling System - 12 Tanks Management
FIXED VERSION - Hard Stop at Minimum Inventory + Cargo Tracking Fixes
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta, date
import os
import random
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
    """Use direct tank datetime values to get start/end times for a given status on a given day"""
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

    if status == 'SUSPENDED':
        suspended_start_dt = tank_object.get('suspended_start_datetime')
        suspended_end_dt = tank_object.get('suspended_end_datetime')

        if suspended_start_dt:
            if is_same_date(suspended_start_dt, current_date_str):
                start_time = safe_format_time(suspended_start_dt)

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
        self.actual_cargo_events = []
        self.berth_status = {
            1: {'occupied': False, 'vessel': None, 'cargo_id': None, 'arrival_time': None},
            2: {'occupied': False, 'vessel': None, 'cargo_id': None, 'arrival_time': None}
        }
        self.next_vessel_id = 1
        self.processing_halted = False # Track if processing has been halted
        
    def track_cargo_status(self, cargo_id, status, berth_id=None, cargo_info=None):
        """Track cargo status with complete information"""
        try:
            # Find existing cargo event
            existing_event = None
            for event in self.actual_cargo_events:
                if event.get('cargo_id') == cargo_id:
                    existing_event = event
                    break
            
            if existing_event:
                # Update existing event
                existing_event['status'] = status
                if berth_id:
                    existing_event['berth_id'] = berth_id
                if cargo_info:
                    # Update with any provided cargo info
                    for key, value in cargo_info.items():
                        if value is not None:
                            existing_event[key] = value
            else:
                # Create new event with complete information
                new_event = {
                    'cargo_id': cargo_id,
                    'status': status,
                    'berth_id': berth_id
                }
                
                # Add cargo info if provided
                if cargo_info:
                    new_event.update(cargo_info)
                else:
                    # Try to get info from cargo schedule
                    for scheduled in self.cargo_schedule:
                        if scheduled.get('cargo_id') == cargo_id:
                            new_event['vessel_name'] = scheduled.get('vessel_name', f"Cargo-{cargo_id:03d}")
                            new_event['type'] = scheduled.get('type', 'Unknown')
                            new_event['size'] = scheduled.get('size', 0)
                            new_event['arrival_datetime'] = scheduled.get('arrival_datetime')
                            new_event['dep_back_datetime'] = scheduled.get('dep_back_datetime')
                            break
                
                # Ensure critical fields are present
                if 'vessel_name' not in new_event:
                    new_event['vessel_name'] = f"Cargo-{cargo_id:03d}"
                if 'type' not in new_event:
                    new_event['type'] = 'Unknown'
                if 'size' not in new_event:
                    new_event['size'] = 0
                    
                self.actual_cargo_events.append(new_event)
            
            return True
            
        except Exception as e:
            print(f"Error tracking cargo status: {e}")
            return False

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
        if dt is None:
            return 'N/A'
        try:
           return dt.strftime("%d/%m/%y %H:%M")
        except (AttributeError, ValueError):
            return 'N/A'

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
        """Calculate when first of last 2 READY tanks will start feeding"""
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

    def _select_optimal_vessel_inventory_driven(self, current_inventory, processing_rate, tank_capacity, num_tanks, available_cargos, params):
        """Enhanced vessel selection that aggressively maintains minimum inventory"""
        if not available_cargos or processing_rate <= 0:
            vessels = list(available_cargos.items())
            return random.choice(vessels) if vessels else None

        # READ USER-DEFINED MINIMUM INVENTORY
        min_inventory_bbl = float(params.get('minInventory', 2000000))
        
        # Calculate days of supply remaining
        days_of_supply = current_inventory / processing_rate if processing_rate > 0 else 999
        
        # Count empty tanks
        empty_tanks_count = sum(1 for t in self.full_tank_details if t['status'] == 'EMPTY')

        # AGGRESSIVE VESSEL SELECTION BASED ON INVENTORY LEVELS
        # Force largest vessels when inventory is critically low
        if current_inventory < min_inventory_bbl * 1.5 or days_of_supply < 10 or empty_tanks_count >= 3:
            # CRITICAL: Force largest available vessel
            vessel_priority = ['vlcc', 'suezmax', 'aframax', 'panamax', 'handymax']
            for vessel_type in vessel_priority:
                if vessel_type in available_cargos:
                    print(f"CRITICAL INVENTORY ({current_inventory:,.0f} bbl, {days_of_supply:.1f} days): Forcing {vessel_type.upper()}")
                    return (vessel_type, available_cargos[vessel_type])
        
        elif current_inventory < min_inventory_bbl * 2.5 or days_of_supply < 20:
            # WARNING: Prefer larger vessels
            if 'vlcc' in available_cargos and random.random() > 0.3: # 70% chance
                return ('vlcc', available_cargos['vlcc'])
            if 'suezmax' in available_cargos and random.random() > 0.4: # 60% chance
                return ('suezmax', available_cargos['suezmax'])
        
        # Normal rotation when inventory is healthy
        vessel_options = list(available_cargos.items())
        return vessel_options[len(self.cargo_schedule) % len(vessel_options)]

    def calculate_next_cargo_timing_improved(self, current_inventory, processing_rate, cargo_info, params,
                                                current_day, last_cargo_arrival_day, total_inventory_capacity):
        """Simple rule: Schedule cargo when 5 tanks empty OR below minimum inventory"""
        
        # READ USER'S MINIMUM INVENTORY SETTING
        min_inventory_bbl = float(params.get('minInventory', 2000000))
        
        # Count empty tanks
        empty_tanks_count = sum(1 for t in self.full_tank_details if t['status'] == 'EMPTY')
        
        # Lead time for cargo to arrive
        journey_days = float(params.get('journeyDays', 10))
        pre_journey_days = float(params.get('preJourneyDays', 1))
        
        # SIMPLE RULES:
        # 1. If 5+ tanks empty → cargo arrives NOW
        if empty_tanks_count >= 5:
            return 0, f"EMERGENCY - {empty_tanks_count} tanks empty, cargo must arrive immediately"
        
        # 2. If below minimum inventory → cargo arrives NOW  
        if current_inventory < min_inventory_bbl:
            return 0, f"CRITICAL - Inventory {current_inventory:,.0f} below minimum {min_inventory_bbl:,.0f}"
        
        # 3. Otherwise, schedule to arrive before hitting minimum
        days_until_minimum = (current_inventory - min_inventory_bbl) / processing_rate if processing_rate > 0 else 999
        
        # Account for journey time - need to depart early
        departure_needed_in = days_until_minimum - journey_days - pre_journey_days
        
        if departure_needed_in <= 0:
            return 0, f"Must depart now to maintain minimum {min_inventory_bbl:,.0f} bbl"
        else:
            return round(departure_needed_in), f"Scheduled to maintain >{min_inventory_bbl:,.0f} bbl minimum"

    def _generate_enhanced_cargo_schedule(self, params, departure_mode='solver', lead_time=None):
        """Enhanced cargo scheduling that aggressively maintains minimum inventory"""
        schedule = []
        
        processing_rate = float(params.get('processingRate', 400000))
        pumping_rate = float(params.get('pumpingRate', 30000))
        pre_journey_days = float(params.get('preJourneyDays', 1))
        journey_days = float(params.get('journeyDays', 10))
        pre_discharge_days = float(params.get('preDischargeDays', 1))
        report_days = int(params.get('schedulingWindow', 70))
        tank_capacity = float(params.get('tankCapacity', 500000))
        num_tanks = int(params.get('numTanks', 12))
        
        # READ USER-DEFINED MINIMUM INVENTORY FROM INPUT FIELD
        MIN_INVENTORY = float(params.get('minInventory', 2000000)) # Reads from user input
        
        # Available vessels
        available_cargos = {}
        if params.get('vlccCapacity', 0) > 0:
            available_cargos['vlcc'] = {'size': float(params.get('vlccCapacity', 2000000)), 'name': 'VLCC'}
        if params.get('suezmaxCapacity', 0) > 0:
            available_cargos['suezmax'] = {'size': float(params.get('suezmaxCapacity', 1000000)), 'name': 'Suezmax'}
        if params.get('aframaxCapacity', 0) > 0:
            available_cargos['aframax'] = {'size': float(params.get('aframaxCapacity', 700000)), 'name': 'Aframax'}
        if params.get('panamaxCapacity', 0) > 0:
            available_cargos['panamax'] = {'size': float(params.get('panamaxCapacity', 450000)), 'name': 'Panamax'}
        if params.get('handymaxCapacity', 0) > 0:
            available_cargos['handymax'] = {'size': float(params.get('handymaxCapacity', 350000)), 'name': 'Handymax'}
        
        if not available_cargos:
            print("WARNING: No cargo types defined")
            return []
        
        processing_start_dt = self._get_processing_start_datetime(params)
        
        # Calculate initial inventory
        total_initial_inventory = 0
        for i in range(1, num_tanks + 1):
            tank_level = float(params.get(f'tank{i}Level', 0))
            dead_bottom = float(params.get(f'deadBottom{i}', 10000))
            available = max(0, tank_level - dead_bottom)
            total_initial_inventory += available
        
        cargo_counter = 1
        scheduled_cargo_ids = set()
        current_inventory = total_initial_inventory
        
        # Track when berths will be free
        berth_free_day = {1: 0, 2: 0}
        all_scheduled_cargos = []
        
        # Count initial empty tanks
        empty_tanks_count = 0
        for i in range(1, num_tanks + 1):
            tank_level = float(params.get(f'tank{i}Level', 0))
            dead_bottom = float(params.get(f'deadBottom{i}', 10000))
            if tank_level <= dead_bottom:
                empty_tanks_count += 1
        
        # ENHANCED: Schedule initial vessels more aggressively if inventory is low
        for berth in [1, 2]:
            # Force VLCC if below minimum OR 5+ tanks empty
            if current_inventory < MIN_INVENTORY or empty_tanks_count >= 5:
                vessel = ('vlcc', available_cargos['vlcc']) if 'vlcc' in available_cargos else max(available_cargos.items(), key=lambda x: x[1]['size'])
            elif 'vlcc' in available_cargos:
                vessel = ('vlcc', available_cargos['vlcc'])
            else:
                vessel = max(available_cargos.items(), key=lambda x: x[1]['size'])
            
            if vessel:
                cargo_type_code, cargo_info = vessel
                
                # If 5+ tanks empty or below minimum → arrive immediately
                if empty_tanks_count >= 5 or current_inventory < MIN_INVENTORY:
                    arrival_day = 1 # Arrive ASAP
                else:
                    if berth == 1:
                        arrival_day = 5
                    else:
                        arrival_day = 8
                
                arrival_date = processing_start_dt + timedelta(days=arrival_day)
                departure_date = arrival_date - timedelta(days=journey_days + pre_journey_days)
                
                pumping_days = cargo_info['size'] / (pumping_rate * 24) if pumping_rate > 0 else 3
                dep_back_date = arrival_date + timedelta(days=pre_discharge_days + pumping_days)
                
                berth_free_day[berth] = arrival_day + pre_discharge_days + pumping_days
                vessel_name = f"{cargo_info['name']}-V{cargo_counter:03d}"
                
                cargo_data = {
                    'cargo_id': cargo_counter,
                    'vessel_name': vessel_name,
                    'type': cargo_info['name'],
                    'size': cargo_info['size'],
                    'dep_port': self._format_datetime_output(departure_date),
                    'arrival': self._format_datetime_output(arrival_date),
                    'dep_back': self._format_datetime_output(dep_back_date),
                    'pumping_days': round(pumping_days, 1),
                    'departure_datetime': departure_date,
                    'arrival_datetime': arrival_date,
                    'dep_back_datetime': dep_back_date,
                    'departure_day': max(1, (departure_date.date() - processing_start_dt.date()).days + 1),
                    'arrival_day': (arrival_date.date() - processing_start_dt.date()).days + 1,
                    'scheduling_reason': f"Berth {berth}: Initial cargo - Maintaining minimum {MIN_INVENTORY:,.0f} bbl",
                    'planned_berth': berth
                }
                
                all_scheduled_cargos.append(cargo_data)
                scheduled_cargo_ids.add(cargo_counter)
                current_inventory += cargo_info['size']
                cargo_counter += 1
        
        # Continue scheduling with enhanced inventory management
        simulation_day = 0
        max_cargos = 200 # Use a high, non-limiting number
        
        while cargo_counter <= max_cargos and simulation_day < report_days + 100:
            # Simulate daily consumption
            current_inventory -= processing_rate
            
            # Check inventory levels
            days_of_supply = current_inventory / processing_rate if processing_rate > 0 else 0
            
            for berth in [1, 2]:
                if cargo_counter > max_cargos:
                    break
                
                berth_becomes_free = berth_free_day[berth]
                
                # SIMPLE RULES: Check if we need emergency cargo
                # Rule 1: If 5+ tanks would be empty → schedule immediate arrival
                # Rule 2: If below minimum inventory → schedule immediate arrival
                empty_tank_projection = sum(1 for t in range(1, num_tanks + 1) if current_inventory < processing_rate * 5)
                
                if empty_tank_projection >= 5 or current_inventory < MIN_INVENTORY:
                    # EMERGENCY - need cargo NOW
                    should_schedule = True
                    next_arrival_day = simulation_day + 1
                else:
                    # Normal scheduling - use berth availability
                    should_schedule = (simulation_day >= berth_becomes_free - journey_days - pre_journey_days - 2)
                    next_arrival_day = max(berth_becomes_free + 4, simulation_day + journey_days + 5)
                    
                if should_schedule:
                    # Select vessel based on simple rules
                    if current_inventory < MIN_INVENTORY or empty_tank_projection >= 5:
                        # EMERGENCY - use largest vessel available
                        vessel = None
                        for v_type in ['vlcc', 'suezmax', 'aframax', 'panamax', 'handymax']:
                            if v_type in available_cargos:
                                vessel = (v_type, available_cargos[v_type])
                                break
                    else:
                        # Normal rotation
                        vessel_options = list(available_cargos.items())
                        vessel = vessel_options[len(all_scheduled_cargos) % len(vessel_options)]
                    
                    if vessel:
                        cargo_type_code, cargo_info = vessel
                        
                        # Calculate when cargo needs to arrive to maintain minimum
                        days_until_critical = (current_inventory - MIN_INVENTORY * 1.25) / processing_rate if processing_rate > 0 else 0
                        
                        # Schedule arrival
                        if days_until_critical < journey_days:
                            # Emergency scheduling
                            next_arrival_day = simulation_day + 1
                        else:
                            # Normal scheduling after berth is free
                            next_arrival_day = max(berth_becomes_free + 2, simulation_day + journey_days + 4)
                        
                        arrival_date = processing_start_dt + timedelta(days=next_arrival_day)
                        departure_date = arrival_date - timedelta(days=journey_days + pre_journey_days)
                        
                        if departure_date >= processing_start_dt:
                            pumping_days = cargo_info['size'] / (pumping_rate * 24) if pumping_rate > 0 else 3
                            dep_back_date = arrival_date + timedelta(days=pre_discharge_days + pumping_days)
                            
                            berth_free_day[berth] = next_arrival_day + pre_discharge_days + pumping_days
                            vessel_name = f"{cargo_info['name']}-V{cargo_counter:03d}"
                            
                            schedule_reason = f"Maintaining >{MIN_INVENTORY:,.0f} bbl minimum"
                            if current_inventory < MIN_INVENTORY:
                                schedule_reason = f"CRITICAL: Below minimum inventory {MIN_INVENTORY:,.0f} bbl"
                            elif empty_tank_projection >= 5:
                                schedule_reason = f"EMERGENCY: {empty_tank_projection} tanks projected empty"
                            
                            cargo_data = {
                                'cargo_id': cargo_counter,
                                'vessel_name': vessel_name,
                                'type': cargo_info['name'],
                                'size': cargo_info['size'],
                                'dep_port': self._format_datetime_output(departure_date),
                                'arrival': self._format_datetime_output(arrival_date),
                                'dep_back': self._format_datetime_output(dep_back_date),
                                'pumping_days': round(pumping_days, 1),
                                'departure_datetime': departure_date,
                                'arrival_datetime': arrival_date,
                                'dep_back_datetime': dep_back_date,
                                'departure_day': max(1, (departure_date.date() - processing_start_dt.date()).days + 1),
                                'arrival_day': (arrival_date.date() - processing_start_dt.date()).days + 1,
                                'scheduling_reason': schedule_reason,
                                'planned_berth': berth
                            }
                            
                            all_scheduled_cargos.append(cargo_data)
                            scheduled_cargo_ids.add(cargo_counter)
                            current_inventory += cargo_info['size']
                            cargo_counter += 1
            
            simulation_day += 1
        
        # Sort and renumber
        all_scheduled_cargos.sort(key=lambda x: x['arrival_datetime'])
        for i, cargo in enumerate(all_scheduled_cargos, 1):
            cargo['cargo_id'] = i
            cargo['vessel_name'] = cargo['vessel_name'].split('-')[0] + f"-V{i:03d}"
        
        return all_scheduled_cargos

    def _check_cargo_arrival(self, current_day, cargo_schedule):
        """Check if any cargo arrives on the given date"""
        try:
            current_date = current_day.date() if hasattr(current_day, 'date') else current_day
        except Exception:
            current_date = current_day
        arrivals = []
        for cargo in cargo_schedule:
            arrival_dt = cargo.get('arrival_datetime')
            if arrival_dt and arrival_dt.date() == current_date:
                arrivals.append({
                    'size': cargo['size'],
                    'type': cargo['type'],
                    'cargo_id': cargo['cargo_id'],
                    'vessel_name': cargo.get('vessel_name', f"{cargo['type']}-{cargo['cargo_id']:03d}"),
                    'arrival_datetime': cargo.get('arrival_datetime'),
                    'dep_back_datetime': cargo.get('dep_back_datetime'),
                    'planned_berth': cargo.get('planned_berth', None) # Pass through planned berth
                })
        return arrivals if arrivals else None

    def _find_earliest_empty_tank(self, tanks, tanks_feeding_today):
        """Find the earliest emptied tank that is not currently feeding"""
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
        """Select earliest filled tank (FIFO)"""
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

    def run_simulation(self, params):
        """Run simulation with HARD STOP at minimum inventory"""
        num_tanks = int(params.get('numTanks', 12))

        # Initialize waiting vessels list
        waiting_vessels = []

        # Initialize simulation state
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
        self.actual_cargo_events = []
        self.berth_status = {
            1: {'occupied': False, 'vessel': None, 'cargo_id': None},
            2: {'occupied': False, 'vessel': None, 'cargo_id': None}
        }
        self.next_vessel_id = 1
        self.processing_halted = False

        try:
            processing_rate = float(params.get('processingRate', 50000))
            if processing_rate <= 0:
                raise ValueError("Processing rate must be greater than 0")
            settling_time_days = float(params.get('settlingTime', 2))
            lab_testing_days = float(params.get('labTestingDays', 1))


            # READ USER-DEFINED MINIMUM INVENTORY
            MIN_INVENTORY = float(params.get('minInventory', 2000000))

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
                'message': f'Simulation started on {processing_start_dt.strftime("%d/%m/%y %H:%M")} with processing rate: {processing_rate:,.0f} bbl/day, Min Inventory: {MIN_INVENTORY:,.0f} bbl (HARD STOP)'
            })

            # Generate cargo schedule with hard constraints
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
                    'was_empty_before_filling': False,
                    'volume_at_day_start': 0,
                    'filling_start_volume': 0,
                    'currently_filling_by_cargo': None # Track which cargo is filling this tank
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
            active_cargos = []
            tanks_emptied_during_day = []

            # Store tanks in full_tank_details for use in other methods
            self.full_tank_details = tanks

            # Run day-by-day simulation
            for day in range(1, report_days + 1):
                current_date = (base_date + timedelta(days=day-1)).date()
                actual_date = processing_start_dt + timedelta(days=day-1)

                display_day = day
                tanks_emptied_during_day = []

                # Track volume at start of day
                for tank in tanks:
                    tank['volume_at_day_start'] = tank['volume']
                    tank['fed_today'] = False
                    tank['emptied_time_today'] = None
                    tank['daily_consumption'] = 0
                    tank['daily_fill_volume'] = 0
                    tank['continuing_fill_tomorrow'] = False
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

                # HARD STOP: Check if below minimum inventory
                if starting_inventory < MIN_INVENTORY and not self.processing_halted:
                    self.processing_halted = True
                    self.alerts.append({
                        'type': 'danger',
                        'day': actual_date.strftime('%d/%m'),
                        'message': f'PROCESSING HALTED: Inventory {starting_inventory:,.0f} bbl BELOW minimum {MIN_INVENTORY:,.0f} bbl'
                    })
                elif starting_inventory >= MIN_INVENTORY and self.processing_halted:
                    self.processing_halted = False
                    self.alerts.append({
                        'type': 'success',
                        'day': actual_date.strftime('%d/%m'),
                        'message': f'PROCESSING RESUMED: Inventory {starting_inventory:,.0f} bbl above minimum {MIN_INVENTORY:,.0f} bbl'
                    })

                # Warning if approaching minimum
                if starting_inventory < MIN_INVENTORY * 1.2 and not self.processing_halted:
                    self.alerts.append({
                        'type': 'warning',
                        'day': actual_date.strftime('%d/%m'),
                        'message': f'WARNING: Inventory {starting_inventory:,.0f} bbl approaching minimum {MIN_INVENTORY:,.0f} bbl'
                    })

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

                # Process tank status transitions
                for tank in tanks:
                    end_of_today = base_date + timedelta(days=day)

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

                            tank['lab_testing_end_datetime'] = settling_end_dt + timedelta(days=lab_testing_days)
                            tank['daily_consumption'] = 0

                            self.alerts.append({
                                'type': 'info', 'day': actual_date.strftime('%d/%m'),
                                'message': f'Tank {tank["id"]} SETTLING complete at {settling_end_dt.strftime("%H:%M")}, starts LAB_TESTING for {lab_testing_days} days until {tank["lab_testing_end_datetime"].strftime("%d/%m %H:%M")}'
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

                # FEEDING LOGIC with HARD STOP at minimum inventory
                if not self.processing_halted:
                    while processing_demand_today > 0:
                        # Check if processing would bring us below minimum
                        current_inventory = sum(t['available'] for t in tanks)
                        if current_inventory - processing_demand_today < MIN_INVENTORY:
                            # Calculate how much we can process before hitting minimum
                            allowed_processing = max(0, current_inventory - MIN_INVENTORY)
                            if allowed_processing == 0:
                                self.processing_halted = True
                                self.alerts.append({
                                    'type': 'danger', 'day': actual_date.strftime('%d/%m'),
                                    'message': f'PROCESSING STOPPED: Would violate minimum inventory {MIN_INVENTORY:,.0f} bbl. {processing_demand_today:,.0f} bbl demand unmet!'
                                })
                                break
                            else:
                                # Partial processing up to minimum
                                processing_demand_today = allowed_processing
                                self.alerts.append({
                                    'type': 'warning', 'day': actual_date.strftime('%d/%m'),
                                    'message': f'PARTIAL PROCESSING: Limited to {allowed_processing:,.0f} bbl to maintain minimum {MIN_INVENTORY:,.0f} bbl'
                                })

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
                                # NO TANKS AVAILABLE
                                self.alerts.append({
                                    'type': 'danger', 'day': actual_date.strftime('%d/%m'),
                                    'message': f'NO TANKS AVAILABLE: Processing stopped with {processing_demand_today:,.0f} bbl demand unmet!'
                                })
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

                # CARGO ARRIVAL with proper berth management
                if not (disruption_duration > 0 and disruption_start <= day < disruption_start + disruption_duration):
                    arrival_infos = self._check_cargo_arrival(current_date, self.cargo_schedule)
                    if arrival_infos:
                        for arrival_info in arrival_infos:
                            # FIXED: Only update cargo type here ONCE
                            day_data['arrivals'] += arrival_info['size']
                            vessel_type_short = arrival_info['type']
                            if day_data['cargo_type']:
                                day_data['cargo_type'] += f"/{vessel_type_short}"
                            else:
                                day_data['cargo_type'] = vessel_type_short

                            planned_berth = arrival_info.get('planned_berth')
                            berth_assigned = None

                            if planned_berth and not self.berth_status[planned_berth]['occupied']:
                                berth_assigned = planned_berth
                            else:
                                for berth_id in [1, 2]:
                                    if not self.berth_status[berth_id]['occupied']:
                                        berth_assigned = berth_id
                                        break

                            if berth_assigned and len(active_cargos) < 2:
                                self.berth_status[berth_assigned]['occupied'] = True
                                self.berth_status[berth_assigned]['vessel'] = arrival_info['vessel_name']
                                self.berth_status[berth_assigned]['cargo_id'] = arrival_info['cargo_id']

                                new_cargo = arrival_info.copy()
                                new_cargo['berth_id'] = berth_assigned
                                new_cargo['remaining_volume'] = new_cargo['size']
                                new_cargo['pumping_start_time'] = new_cargo['arrival_datetime'] + timedelta(days=float(self.initial_params.get('preDischargeDays', 1)))
                                active_cargos.append(new_cargo)
                                
                                # Track with complete info
                                cargo_info = {
                                    'vessel_name': new_cargo['vessel_name'],
                                    'type': new_cargo['type'],
                                    'size': new_cargo['size'],
                                    'actual_arrival': new_cargo['arrival_datetime'],
                                    'actual_pumping_start': new_cargo['pumping_start_time'],
                                    'actual_pumping_end': None,
                                    'actual_departure': None
                                }
                                self.track_cargo_status(arrival_info['cargo_id'], 'ARRIVED', berth_assigned, cargo_info)
                                
                                # REMOVED: Don't update day_data here - already done above
                                    
                                self.alerts.append({
                                    'type': 'success', 'day': actual_date.strftime('%d/%m'),
                                    'message': f"BERTH {berth_assigned}: {new_cargo['vessel_name']} arrived at {new_cargo['arrival_datetime'].strftime('%H:%M')}. Cargo: {new_cargo['size']:,.0f} bbl"
                                })
                            else:
                                waiting_vessels.append(arrival_info)
                                self.alerts.append({
                                    'type': 'warning', 'day': actual_date.strftime('%d/%m'),
                                    'message': f"Both berths occupied. {arrival_info['vessel_name']} waiting at anchorage."
                                })

                # CARGO FILLING LOGIC
                total_cargo_opening_stock = 0
                total_cargo_consumption_today = 0
                total_cargo_closing_stock = 0
                cargos_to_remove = []

                for cargo_idx, active_cargo in enumerate(active_cargos):
                    if active_cargo['remaining_volume'] > 0 and current_date >= active_cargo['pumping_start_time'].date():
                        cargo_opening_stock = active_cargo['remaining_volume']
                        cargo_consumption_today = 0

                        if active_cargo['remaining_volume'] == active_cargo['size']:
                            # Update tracking when pumping starts
                            for event in self.actual_cargo_events:
                                if event['cargo_id'] == active_cargo['cargo_id']:
                                    event['status'] = 'PUMPING'
                                    break

                        pumping_start_this_day = max(datetime.combine(current_date, datetime.min.time()), active_cargo['pumping_start_time'])
                        hours_to_pump_today = (day_end_time - pumping_start_this_day).total_seconds() / 3600
                        volume_to_pump_today = min(hours_to_pump_today * pumping_rate_per_hour, active_cargo['remaining_volume'])
                        current_pumping_time = pumping_start_this_day
                        tanks_available_for_filling = False

                        while volume_to_pump_today > 0 and active_cargo['remaining_volume'] > 0:
                            target_tank = next((t for t in tanks if t['status'] == 'FILLING' and t.get('filling_cargo_id') == active_cargo['cargo_id'] and t.get('currently_filling_by_cargo') == active_cargo['cargo_id']), None)
                            
                            if not target_tank:
                                suspended_tanks = [t for t in tanks if t['status'] == 'SUSPENDED' and t.get('currently_filling_by_cargo') is None]
                                if suspended_tanks:
                                    target_tank = suspended_tanks[0]
                                    target_tank['filling_start_volume'] = target_tank['volume']
                                    target_tank['was_empty_before_filling'] = False
                                    tanks_available_for_filling = True
                                    self.alerts.append({'type': 'info', 'day': actual_date.strftime('%d/%m'), 'message': f"Resuming fill of SUSPENDED Tank {target_tank['id']} with {active_cargo['vessel_name']}"})
                                elif not target_tank:
                                    for emptied_info in tanks_emptied_during_day:
                                        if pumping_start_this_day >= emptied_info['time']:
                                            potential_tank = next((t for t in tanks if t['id'] == emptied_info['tank_id'] and t['status'] == 'EMPTY' and t.get('currently_filling_by_cargo') is None), None)
                                            if potential_tank and potential_tank['id'] not in tanks_feeding_today and not potential_tank['fed_today']:
                                                target_tank = potential_tank
                                                target_tank['filling_start_volume'] = target_tank['volume']
                                                target_tank['was_empty_before_filling'] = True
                                                tanks_available_for_filling = True
                                                break
                                if not target_tank:
                                    empty_tanks = [t for t in tanks if t['status'] == 'EMPTY' and t['id'] not in tanks_feeding_today and not t['fed_today'] and t.get('currently_filling_by_cargo') is None]
                                    if empty_tanks:
                                        target_tank = empty_tanks[0]
                                        target_tank['filling_start_volume'] = target_tank['volume']
                                        target_tank['was_empty_before_filling'] = True
                                        tanks_available_for_filling = True
                            else:
                                tanks_available_for_filling = True
                            
                            if target_tank:
                                target_tank['currently_filling_by_cargo'] = active_cargo['cargo_id']
                                if target_tank['status'] != 'FILLING':
                                    start_fill_time = current_pumping_time
                                    if not target_tank.get('filling_start_volume'):
                                        target_tank['filling_start_volume'] = target_tank['volume']
                                    target_tank['filling_start_datetime'] = start_fill_time
                                    target_tank['filling_cargo_id'] = active_cargo['cargo_id']
                                    target_tank['vessel_arrival_datetime'] = active_cargo.get('arrival_datetime')
                                    target_tank['vessel_dep_datetime'] = active_cargo.get('dep_back_datetime')
                                    self.filling_events_log.append({'tank_id': target_tank['id'], 'start': start_fill_time, 'end': None, 'settle_start': None, 'lab_start': None, 'ready_time': None, 'cargo_type': active_cargo['vessel_name']})
                                    self.alerts.append({'type': 'info', 'day': actual_date.strftime('%d/%m'), 'message': f"BERTH {active_cargo.get('berth_id', '?')}: Filling Tank {target_tank['id']} from {active_cargo['vessel_name']} at {start_fill_time.strftime('%H:%M')}"})
                            else:
                                cargo_departed = False
                                if active_cargo.get('dep_back_datetime') and active_cargo['dep_back_datetime'].date() <= current_date:
                                    cargo_departed = True
                                if not cargo_departed and not tanks_available_for_filling:
                                    self.alerts.append({'type': 'danger', 'day': actual_date.strftime('%d/%m'), 'message': f"DEMURRAGE: {active_cargo['vessel_name']} (Berth {active_cargo.get('berth_id', '?')}) - no empty tank"})
                                break

                            space_in_tank = tank_capacity - target_tank['volume']
                            volume_for_this_tank = min(space_in_tank, volume_to_pump_today, active_cargo['remaining_volume'])
                            
                            if volume_for_this_tank > 0:
                                self.daily_discharge_log.append({'date': actual_date.strftime('%d/%m/%y'), 'cargo_type': active_cargo['vessel_name'], 'tank_id': target_tank['id'], 'volume_filled': volume_for_this_tank})
                                target_tank['status'] = 'FILLING'
                                target_tank['volume'] += volume_for_this_tank
                                target_tank['daily_fill_volume'] += volume_for_this_tank
                                target_tank['daily_consumption'] = -target_tank['daily_fill_volume']
                                active_cargo['remaining_volume'] -= volume_for_this_tank
                                volume_to_pump_today -= volume_for_this_tank
                                cargo_consumption_today += volume_for_this_tank
                                pumping_hours = volume_for_this_tank / pumping_rate_per_hour if pumping_rate_per_hour > 0 else 0
                                current_pumping_time += timedelta(hours=pumping_hours)
                                
                                if target_tank['volume'] >= tank_capacity - 1:
                                    target_tank['volume'] = tank_capacity
                                    filling_end_time = current_pumping_time
                                    target_tank['filling_end_datetime'] = filling_end_time
                                    target_tank['currently_filling_by_cargo'] = None
                                    target_tank['status'] = 'FILLED'
                                    target_tank['filled_datetime'] = filling_end_time
                                    target_tank['daily_consumption'] = 0
                                    target_tank['status'] = 'SETTLING'
                                    target_tank['settling_start_datetime'] = filling_end_time

                                    target_tank['settling_end_datetime'] = filling_end_time + timedelta(days=settling_time_days)
                                    self.alerts.append({'type': 'info', 'day': actual_date.strftime('%d/%m'), 'message': f"Tank {target_tank['id']} FILLED at {filling_end_time.strftime('%H:%M')} with {active_cargo['vessel_name']}, starts SETTLING for {settling_time_days} days"})

                            if active_cargo['remaining_volume'] <= 0:
                                actual_pumping_end_time = current_pumping_time
                                filling_end_time = current_pumping_time
                                if target_tank:
                                    target_tank['filling_end_datetime'] = filling_end_time
                                    target_tank['currently_filling_by_cargo'] = None
                                
                                # Update tracking when pumping completes
                                for event in self.actual_cargo_events:
                                    if event['cargo_id'] == active_cargo.get('cargo_id'):
                                        event['actual_pumping_end'] = actual_pumping_end_time
                                        event['actual_departure'] = actual_pumping_end_time
                                        event['status'] = 'COMPLETED'
                                        break
                                        
                                if target_tank and (target_tank['volume'] > target_tank['dead_bottom'] and target_tank['volume'] < tank_capacity):
                                    filling_start_dt = target_tank.get('filling_start_datetime')
                                    if filling_start_dt:
                                        hours_pumped = (filling_end_time - filling_start_dt).total_seconds() / 3600
                                        volume_pumped = hours_pumped * pumping_rate_per_hour
                                        suspended_volume = target_tank.get('filling_start_volume', target_tank['dead_bottom']) + volume_pumped
                                        target_tank['suspended_volume'] = suspended_volume
                                    else:
                                        target_tank['suspended_volume'] = target_tank['volume']
                                    target_tank['status'] = 'SUSPENDED'
                                    target_tank['suspended_start_datetime'] = filling_end_time
                                    target_tank['suspended_end_datetime'] = filling_end_time + timedelta(hours=1)
                                    target_tank['daily_consumption'] = 0
                                    self.alerts.append({'type': 'warning', 'day': actual_date.strftime('%d/%m'), 'message': f"Tank {target_tank['id']} SUSPENDED at {filling_end_time.strftime('%H:%M')}. Volume: {target_tank['suspended_volume']:,.0f} bbl"})
                        
                        cargo_closing_stock = active_cargo['remaining_volume']
                        total_cargo_opening_stock += cargo_opening_stock
                        total_cargo_consumption_today += cargo_consumption_today
                        total_cargo_closing_stock += cargo_closing_stock

                        if active_cargo['remaining_volume'] <= 0:
                            berth_id = active_cargo.get('berth_id', 1)
                            self.alerts.append({'type': 'success', 'day': actual_date.strftime('%d/%m'), 'message': f"BERTH {berth_id}: {active_cargo['vessel_name']} completed discharge. Berth now available."})
                            self.berth_status[berth_id]['occupied'] = False
                            self.berth_status[berth_id]['vessel'] = None
                            self.berth_status[berth_id]['cargo_id'] = None
                            cargos_to_remove.append(cargo_idx)

                            # Check if there are waiting vessels for this freed berth
                            if waiting_vessels:
                                next_vessel = waiting_vessels.pop(0)
                                
                                self.berth_status[berth_id]['occupied'] = True
                                self.berth_status[berth_id]['vessel'] = next_vessel['vessel_name']
                                self.berth_status[berth_id]['cargo_id'] = next_vessel['cargo_id']
                                
                                new_cargo = next_vessel.copy()
                                new_cargo['berth_id'] = berth_id
                                new_cargo['remaining_volume'] = new_cargo['size']
                                new_cargo['pumping_start_time'] = datetime.combine(current_date, datetime.now().time()) + timedelta(days=float(self.initial_params.get('preDischargeDays', 1)))
                                active_cargos.append(new_cargo)
                                
                                # Track the waiting vessel now arriving with complete info
                                cargo_info = {
                                    'vessel_name': new_cargo['vessel_name'],
                                    'type': new_cargo['type'],
                                    'size': new_cargo['size'],
                                    'actual_arrival': datetime.combine(current_date, datetime.now().time()),
                                    'actual_pumping_start': new_cargo['pumping_start_time'],
                                    'actual_pumping_end': None,
                                    'actual_departure': None
                                }
                                self.track_cargo_status(next_vessel['cargo_id'], 'ARRIVED', berth_id, cargo_info)
                                
                                # REMOVED: Don't add to arrivals again - waiting vessels were already counted
                                
                                self.alerts.append({
                                    'type': 'success',
                                    'day': actual_date.strftime('%d/%m'),
                                    'message': f"BERTH {berth_id}: Assigned waiting vessel {next_vessel['vessel_name']} at {datetime.now().strftime('%H:%M')}"
                                })

                # Remove completed cargos
                for idx in reversed(cargos_to_remove):
                    active_cargos.pop(idx)
                
                for tank in tanks:
                    if tank.get('currently_filling_by_cargo') and tank['status'] != 'FILLING':
                        tank['currently_filling_by_cargo'] = None
                
                day_data.update({'cargo_opening_stock': total_cargo_opening_stock, 'cargo_consumption_today': total_cargo_consumption_today, 'cargo_closing_stock': total_cargo_closing_stock})
                ending_inventory = sum(t['available'] for t in tanks)
                day_data['end_inventory'] = ending_inventory
                total_usable_capacity = sum(tank_capacity for t in tanks)
                day_data['tank_utilization'] = (ending_inventory / total_usable_capacity) * 100 if total_usable_capacity > 0 else 0

                for tank in tanks:
                    if tank['status'] == 'SUSPENDED' and tank.get('suspended_volume') is not None:
                        closing_stock = tank['suspended_volume']
                        if tank.get('suspended_start_datetime') and tank['suspended_start_datetime'].date() == current_date:
                            opening_stock = tank.get('filling_start_volume', tank['volume_at_day_start'])
                        else:
                            opening_stock = tank.get('suspended_volume', tank['volume_at_day_start'])
                    elif tank['status'] == 'FILLING':
                        if tank.get('filling_start_datetime') and tank['filling_start_datetime'].date() == current_date:
                            if tank.get('was_empty_before_filling'):
                                opening_stock = tank.get('filling_start_volume', tank['dead_bottom'])
                                closing_stock = tank['volume']
                            else:
                                opening_stock = tank.get('filling_start_volume', tank['volume'] - tank['daily_fill_volume'])
                                closing_stock = tank['volume']
                        else:
                            opening_stock = tank['volume_at_day_start']
                            closing_stock = tank['volume']
                    else:
                        opening_stock = tank['volume'] + tank['daily_consumption'] - tank['daily_fill_volume']
                        closing_stock = tank['volume']

                    day_data.update({f'tank{tank["id"]}_level': tank['volume'], f'tank{tank["id"]}_status': tank['status'], f'tank{tank["id"]}_consumption': tank['daily_consumption'], f'tank{tank["id"]}_opening_stock': opening_stock, f'tank{tank["id"]}_closing_stock': closing_stock, f'tank{tank["id"]}_status_start_time': '', f'tank{tank["id"]}_status_end_time': '', f'tank{tank["id"]}_filling_cargo': tank.get('filling_cargo_id', ''), f'tank{tank["id"]}_filled_time': '', f'tank{tank["id"]}_suspended_start': '', f'tank{tank["id"]}_suspended_end': ''})
                    start_time, end_time = populate_tank_times(tank['status'], tank['id'], day_data, self.feeding_events_log, self.filling_events_log, tank)
                    day_data[f'tank{tank["id"]}_status_start_time'] = start_time
                    day_data[f'tank{tank["id"]}_status_end_time'] = end_time
                    if tank['status'] == 'SUSPENDED':
                        if tank.get('suspended_start_datetime') and tank['suspended_start_datetime'].date() == current_date:
                            day_data[f'tank{tank["id"]}_suspended_start'] = tank['suspended_start_datetime'].strftime('%H:%M')
                        if tank.get('suspended_end_datetime') and tank.get('suspended_end_datetime').date() == current_date:
                            day_data[f'tank{tank["id"]}_suspended_end'] = tank['suspended_end_datetime'].strftime('%H:%M')
                    if tank['status'] == 'SETTLING':
                        if tank.get('settling_start_datetime') and tank['settling_start_datetime'].date() == current_date:
                            day_data[f'tank{tank["id"]}_filled_time'] = tank['settling_start_datetime'].strftime('%H:%M')
                
                self.simulation_data.append(day_data)

            self.full_tank_details = tanks
            metrics = self._calculate_metrics(params)
            buffer_info = self._calculate_buffer_stock(params)
            cargo_report = self._generate_cargo_report(params)

            final_feeding_end_dt = None
            for tank in tanks:
                if tank.get('feeding_end_datetime') and (final_feeding_end_dt is None or tank['feeding_end_datetime'] > final_feeding_end_dt):
                    final_feeding_end_dt = tank['feeding_end_datetime']
            
            first_feeding_start_dt = None
            for tank in tanks:
                if tank.get('original_feeding_start') and (first_feeding_start_dt is None or tank['original_feeding_start'] < first_feeding_start_dt):
                    first_feeding_start_dt = tank['original_feeding_start']

            first_filling_start_dt = None
            last_filling_end_dt = None
            for tank in tanks:
                if tank.get('filling_start_datetime') and (first_filling_start_dt is None or tank['filling_start_datetime'] < first_filling_start_dt):
                    first_filling_start_dt = tank['filling_start_datetime']
                if tank.get('filling_end_datetime') and (last_filling_end_dt is None or tank['filling_end_datetime'] > last_filling_end_dt):
                    last_filling_end_dt = tank['filling_end_datetime']

            initial_start_time_str = self._format_datetime_output(first_feeding_start_dt) if first_feeding_start_dt else "N/A"
            final_end_time_str = self._format_datetime_output(final_feeding_end_dt) if final_feeding_end_dt else "N/A"
            first_filling_start_str = self._format_datetime_output(first_filling_start_dt) if first_filling_start_dt else "N/A"
            last_filling_end_str = self._format_datetime_output(last_filling_end_dt) if last_filling_end_dt else "N/A"

            return {'parameters': params, 'simulation_data': self.simulation_data, 'alerts': self.alerts, 'metrics': metrics, 'cargo_schedule': cargo_report,'cargo_report': cargo_report, 'feeding_events_log': self.feeding_events_log, 'filling_events_log': self.filling_events_log, 'daily_discharge_log': self.daily_discharge_log, 'buffer_info': buffer_info, 'initial_start_time': initial_start_time_str, 'final_end_time': final_end_time_str, 'first_filling_start_time': first_filling_start_str, 'last_filling_end_time': last_filling_end_str, 'full_tank_details': self.full_tank_details}
        
        except ZeroDivisionError as e:
            return {'error': f'Division by zero error: {str(e)}. Please check input parameters'}
        except Exception as e:
            import traceback
            traceback.print_exc()
            return {'error': str(e)}

    def _calculate_metrics(self, params):
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

    def _generate_cargo_report(self, params):
        """Generate cargo report with proper data extraction from simulation - FIXED VERSION"""
        cargo_report = []
        
        # Helper function to safely format datetime with None checking
        def safe_format_datetime(dt_value):
            """Safely format datetime, handling None values"""
            if dt_value is None:
                return 'N/A'
            try:
                if hasattr(dt_value, 'strftime'):
                    return dt_value.strftime("%d/%m/%y %H:%M")
                else:
                    return 'N/A'
            except (AttributeError, ValueError):
                return 'N/A'
        
        # Use actual events if available
        if hasattr(self, 'cargo_schedule') and self.cargo_schedule:
            pre_journey_days = float(params.get('preJourneyDays', 1))

            # --- FIX: Define the report window cutoff date ---
            try:
                report_days = int(params.get('schedulingWindow', 70))
                # This method exists in your utils.py file to parse the start date
                start_dt = self._get_processing_start_datetime(params)
                end_date_cutoff = (start_dt + timedelta(days=report_days)).date()
            except Exception:
                end_date_cutoff = None # If it fails, show all cargoes
            # --- END FIX ---

            for cargo in self.cargo_schedule:
                try:
                    arrival_dt = cargo.get('arrival_datetime')

                    # --- FIX: Filter out cargoes arriving after the report window ---
                    if end_date_cutoff and (not arrival_dt or arrival_dt.date() > end_date_cutoff):
                        continue # Skip this cargo
                    # --- END FIX ---
                    
                    departure_dt = cargo.get('departure_datetime')
                    dep_unload_port_dt = cargo.get('dep_back_datetime')
                    cargo_size = cargo.get('size', 0)
                
                    # Check if this cargo actually arrived (from actual_cargo_events)
                    actual_times = None
                    if hasattr(self, 'actual_cargo_events'):
                        for event in self.actual_cargo_events:
                            if event.get('cargo_id') == cargo.get('cargo_id'):
                                actual_times = event
                                break
                
                    # Use actual times if available, otherwise use scheduled
                    status = "Scheduled"
                    # Use actual times and status if available, otherwise use scheduled
                    if actual_times:
                        arrival_dt = actual_times.get('actual_arrival', arrival_dt)
                    
                        if actual_times.get('actual_departure') is not None:
                            dep_unload_port_dt = actual_times['actual_departure']

                        berth_id = actual_times.get('berth_id', cargo.get('planned_berth', 'N/A'))
                        cargo_size = actual_times.get('size', cargo_size)
                        cargo['type'] = actual_times.get('type', cargo.get('type'))
                    
                        # Get the dynamic status from the event log and capitalize it.
                        status = actual_times.get('status', 'Scheduled').title()
                    else:
                        berth_id = cargo.get('planned_berth', 'N/A')
                
                    # Calculate load port time
                    load_port_time_dt = None
                    if departure_dt is not None:
                        try:
                            load_port_time_dt = departure_dt - timedelta(days=pre_journey_days)
                        except Exception:
                            load_port_time_dt = None

                    cargo_report.append({
                        'cargo_id': cargo.get('cargo_id'),
                        'berth': f"BERTH {berth_id}",
                        'vessel_name': cargo.get('vessel_name', 'N/A'), 
                        'type': cargo.get('type', 'Unknown'),
                        'load_port_time': safe_format_datetime(load_port_time_dt),
                        'dep_time': safe_format_datetime(departure_dt),
                        'arrival_time': safe_format_datetime(arrival_dt),
                        'dep_unload_port': safe_format_datetime(dep_unload_port_dt),
                        'size': cargo_size,
                        'status': status, # 
                        'pumping_days': cargo.get('pumping_days', 0), # <-- ADD THIS LINE
                        
                        # FIX for undefined issue: Add old keys for HTML page compatibility
                        'arrival': safe_format_datetime(arrival_dt),
                        'dep_back': safe_format_datetime(dep_unload_port_dt),
                        'dep_port': safe_format_datetime(departure_dt)
                    })
                except Exception as e:
                    print(f"Error processing cargo {cargo.get('cargo_id', 'unknown')}: {e}")
                    continue
       
        return cargo_report