import win32print

def get_printer_queue(printer_name):
    try:
        # Get a handle to the printer
        printer_handle = win32print.OpenPrinter(printer_name)
        
        # Get the print queue information
        queue_info = win32print.GetPrinter(printer_handle, 2)

        # Display the print queue information
        print("Printer Queue Information:")
        print("Printer Name:", queue_info['pPrinterName'])
        print("Driver Name:", queue_info['pDriverName'])
        print("Port Name:", queue_info['pPortName'])
        print("Print Jobs in Queue:", queue_info['cJobs'])
        print("\nPrint Job Information:")

        # Enumerate print jobs in the queue
        for job_id in range(queue_info['cJobs']):
            job_info = win32print.EnumJobs(printer_handle, 0, queue_info['cJobs'], 2)

            # Display information about each print job
            print(f"\nPrint Job #{job_id + 1}:")
            print("Document Name:", job_info[job_id]['pDocument']
                  if job_info[job_id]['pDocument'] else "N/A")
            
            # Get the status code
            status_code = job_info[job_id]['Status']
            
            # Map status codes to human-readable statements
            status_messages = {
                1: 'JOB_STATUS_PAUSED',
                2: 'JOB_STATUS_ERROR',
                4: 'JOB_STATUS_DELETING',
                8: 'JOB_STATUS_SPOOLING',
                16: 'JOB_STATUS_PRINTING',
                32: 'JOB_STATUS_OFFLINE',
                64: 'JOB_STATUS_PAPEROUT',
                128: 'JOB_STATUS_PRINTED',
                256: 'JOB_STATUS_DELETED',
                512: 'JOB_STATUS_BLOCKED_DEVQ',
                1024: 'JOB_STATUS_USER_INTERVENTION',
                2048: 'JOB_STATUS_RESTART',
                4096: 'JOB_STATUS_COMPLETE',
                8208: 'JOB_STATUS_PRINTING'
            }
            
            # Display a human-readable status message
            status_message = status_messages.get(status_code, f'Unknown Status: {status_code}')
            print("Status:", status_message)
            print("Total Pages:", job_info[job_id]['TotalPages'])

    except Exception as e:
        print(f"Error: {e}")

    finally:
        # Close the printer handle
        win32print.ClosePrinter(printer_handle)

# Specify the printer name
printer_name = "MJ8330 thermal printer test"

# Get and display the printer and print job information
get_printer_queue(printer_name)
