import sys
# Make sure we don't start the threads on import by mocking threading.Thread.start if needed, or just let them start.
from app import _load_ag_data_internal, _state

ok, err = _load_ag_data_internal()
print("Success:", ok, "Error:", err)
if ok:
    store_data = _state.get("ag_store_data", [])
    print("Total store records:", len(store_data))
    if store_data:
        print("First record keys:", store_data[0].keys())
        print("First record CM:", store_data[0].get("CM"))
        
        # Check how many have CM not equal to 'Unmapped'
        unmapped_cnt = sum(1 for s in store_data if s.get("CM") == "Unmapped")
        mapped_cnt = sum(1 for s in store_data if s.get("CM") and s.get("CM") != "Unmapped")
        none_cnt = sum(1 for s in store_data if s.get("CM") is None)
        print(f"Mapped: {mapped_cnt}, Unmapped: {unmapped_cnt}, None: {none_cnt}")
