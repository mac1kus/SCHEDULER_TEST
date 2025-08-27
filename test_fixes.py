#!/usr/bin/env python3
"""
Test script to verify the processing logic fixes
"""

import json
import sys
import os

# Add the current directory to Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import AdvancedRefineryCrudeScheduler

def test_processing_logic():
    """Test the processing logic with the scenario described by the user"""
    
    # Create test parameters similar to the user's scenario
    test_params = {
        'processingRate': 200000,  # 200,000 bbl/day as in last_inputs.json
        'tankCapacity': 600000,
        'pumpingRate': 30000,
        'departureMode': 'solver',
        'preJourneyDays': 1,
        'journeyDays': 10,
        'preDischargeDays': 1,
        'settlingTime': 2,
        'labTestingDays': 1,
        'bufferDays': 2,
        'schedulingWindow': 65,
        'disruptionDuration': 0,
        'disruptionStart': 20,
        'vlccCapacity': 2000000,
        'suezmaxCapacity': 1000000,
        'aframaxCapacity': 750000,
        'panamaxCapacity': 0,
        'handymaxCapacity': 0,
        'defaultDeadBottom': 10000,
        'bufferVolume': 500
    }
    
    # Set tank levels - all tanks at capacity initially
    for i in range(1, 13):
        test_params[f'tank{i}Level'] = 600000
        test_params[f'deadBottom{i}'] = 10000
    
    # Create scheduler and run simulation
    scheduler = AdvancedRefineryCrudeScheduler()
    results = scheduler.run_simulation(test_params)
    
    if 'error' in results:
        print(f"❌ Simulation failed: {results['error']}")
        return False
    
    # Check for the specific issues mentioned by the user
    simulation_data = results['simulation_data']
    alerts = results['alerts']
    
    print("🔍 Testing Processing Logic Fixes...")
    print(f"📊 Processing Rate: {test_params['processingRate']:,} bbl/day")
    print(f"📅 Simulation Days: {len(simulation_data)}")
    
    # Check for incomplete processing days
    incomplete_days = []
    mass_balance_errors = []
    
    for day_data in simulation_data:
        day = day_data['day']
        processing = day_data['processing']
        target_processing = test_params['processingRate']
        
        # Check for incomplete processing
        if processing < target_processing:
            incomplete_days.append({
                'day': day,
                'processing': processing,
                'target': target_processing,
                'shortfall': target_processing - processing
            })
        
        # Check mass balance
        total_consumption = sum(day_data.get(f'tank{i}_consumption', 0) for i in range(1, 13))
        if abs(total_consumption - processing) > 0.01:
            mass_balance_errors.append({
                'day': day,
                'consumption': total_consumption,
                'processing': processing,
                'difference': abs(total_consumption - processing)
            })
    
    # Check for specific alerts about processing issues
    processing_alerts = [alert for alert in alerts if 'PROCESSING' in alert['message'] or 'INCOMPLETE' in alert['message']]
    
    print(f"\n📋 Results:")
    print(f"   • Incomplete processing days: {len(incomplete_days)}")
    print(f"   • Mass balance errors: {len(mass_balance_errors)}")
    print(f"   • Processing-related alerts: {len(processing_alerts)}")
    
    if incomplete_days:
        print(f"\n⚠️  Incomplete Processing Days:")
        for issue in incomplete_days[:5]:  # Show first 5
            print(f"   Day {issue['day']}: {issue['processing']:,} of {issue['target']:,} bbl ({issue['shortfall']:,} shortfall)")
    
    if mass_balance_errors:
        print(f"\n❌ Mass Balance Errors:")
        for error in mass_balance_errors[:5]:  # Show first 5
            print(f"   Day {error['day']}: Consumption {error['consumption']:,} ≠ Processing {error['processing']:,}")
    
    if processing_alerts:
        print(f"\n🔔 Processing Alerts:")
        for alert in processing_alerts[:5]:  # Show first 5
            print(f"   Day {alert['day']}: {alert['message']}")
    
    # Check if the fixes are working
    success = True
    
    if incomplete_days:
        print(f"\n❌ Issue: Found {len(incomplete_days)} days with incomplete processing")
        success = False
    else:
        print(f"\n✅ Fix 1: No incomplete processing days detected")
    
    if mass_balance_errors:
        print(f"\n❌ Issue: Found {len(mass_balance_errors)} mass balance errors")
        success = False
    else:
        print(f"\n✅ Fix 2: No mass balance errors detected")
    
    if processing_alerts:
        print(f"\n✅ Fix 3: Processing alerts are being generated to notify about issues")
    else:
        print(f"\n✅ Fix 3: No processing issues detected")
    
    return success

if __name__ == "__main__":
    print("🧪 Testing Refinery Processing Logic Fixes")
    print("=" * 50)
    
    success = test_processing_logic()
    
    print("\n" + "=" * 50)
    if success:
        print("🎉 All tests passed! Processing logic fixes are working correctly.")
    else:
        print("⚠️  Some issues detected. Please review the results above.")
    
    print("\n💡 To test with the web interface:")
    print("   1. Run: python app.py")
    print("   2. Open: http://localhost:5000")
    print("   3. Load saved inputs and run simulation")
    print("   4. Check alerts for processing issues")
