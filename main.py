import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime, timedelta
from evtx import PyEvtxParser
from bs4 import BeautifulSoup

class UptimeTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title('Windows Uptime Tracker')
        self.create_widgets()
        self.prefill_dates()


    def create_widgets(self):
        tk.Label(self.root, text='Start Date (mm/dd/yyyy):').grid(row=0, column=0)
        self.start_date_entry = tk.Entry(self.root)
        self.start_date_entry.grid(row=0, column=1)

        tk.Label(self.root, text='End Date (mm/dd/yyyy):').grid(row=1, column=0)
        self.end_date_entry = tk.Entry(self.root)
        self.end_date_entry.grid(row=1, column=1)

        track_button = tk.Button(self.root, text='Track Uptime', command=self.track_uptime)
        track_button.grid(row=2, column=0, columnspan=2)

        export_button = tk.Button(self.root, text='Export to CSV', command=self.export_to_csv)
        export_button.grid(row=3, column=0, columnspan=2)

    def track_uptime(self):
        start_date = self.start_date_entry.get()
        end_date = self.end_date_entry.get()
        uptime_dict = self.get_uptime(start_date, end_date)
        self.uptime_df = pd.DataFrame(list(uptime_dict.items()), columns=['Date', 'Uptime'])
        # Check if the 'Uptime' column contains timedelta objects and convert them to seconds
        self.uptime_df['Uptime'] = self.uptime_df['Uptime'].apply(lambda x: x.total_seconds() if isinstance(x, timedelta) else x)

        # Now convert seconds to hours
        self.uptime_df['Uptime'] = self.uptime_df['Uptime'] / 3600

        self.plot_uptime()


    def get_uptime(self, start_date, end_date):
        """
        Parses the Windows Event Log to calculate the system uptime between two dates.

        Parameters:
        start_date (str): The start date in 'mm/dd/yyyy' format.
        end_date (str): The end date in 'mm/dd/yyyy' format.

        Returns:
        dict: A dictionary with dates as keys and uptime in seconds as values.
        """
        start_date = datetime.strptime(start_date, '%m/%d/%Y').date()
        end_date = datetime.strptime(end_date, '%m/%d/%Y').date()
        uptime_dict = {}
        sessions_dict = {}
        log_file = r'system.evtx'  # Path to the system event log file
        parser = PyEvtxParser(log_file)

        # Store start and end times of sessions
        sessions = []
        last_event_time = None  # Keep track of the last event time

        for record in parser.records():
            event_data = record['data']
            soup = BeautifulSoup(event_data, 'lxml-xml')
            event_id = int(soup.find('EventID').text)
            event_time_str = soup.find('TimeCreated')['SystemTime']
            event_time = datetime.strptime(event_time_str, '%Y-%m-%dT%H:%M:%S.%fZ')
            event_date = event_time.date()

            if start_date <= event_date <= end_date:
                last_event_time = event_time  # Update last event time

                if event_id == 7001:  # System start event
                    # If there's an open session without an end time, close it with the last event time
                    if sessions and sessions[-1]['end'] is None:
                        if last_event_time > sessions[-1]['start']:
                            sessions[-1]['end'] = last_event_time
                        else:
                            print(f"Error: End time {last_event_time} is before start time {sessions[-1]['start']}")
                            continue
                    sessions.append({'start': event_time, 'end': None})

                elif event_id == 6006 or event_id == 6008:  # Normal or unexpected shutdown event
                    # Find the last session without an end time and set its end time
                    if sessions and sessions[-1]['end'] is None:
                        if event_time > sessions[-1]['start']:
                            sessions[-1]['end'] = event_time
                        else:
                            print(f"Error: End time {event_time} is before start time {sessions[-1]['start']}")
                            continue

        # Handle the case where the last recorded event is a start event without an end event
        if sessions and sessions[-1]['end'] is None:
            if last_event_time > sessions[-1]['start']:
                sessions[-1]['end'] = last_event_time
            else:
                print(f"Error: Last event time {last_event_time} is before start time {sessions[-1]['start']}")

        # Calculate uptime for each session and sum for each day
        for session in sessions:
            if session['end'] is not None:
                session_duration = session['end'] - session['start']
                session_date = session['start'].date()
                if session_date not in uptime_dict:
                    uptime_dict[session_date] = timedelta()
                uptime_dict[session_date] += session_duration
                if session_date not in sessions_dict:
                    sessions_dict[session_date] = []
                sessions_dict[session_date].append(session_duration)

        # Convert timedeltas to seconds and sum up the session durations
        for date in uptime_dict:
            uptime_dict[date] = sum((session.total_seconds() for session in sessions_dict[date]), 0)

        print("Uptime Dictionary:", uptime_dict)
        print("Sessions Dictionary:", sessions_dict)
        print("Sessions:", sessions)
        return uptime_dict


    def plot_uptime(self):
        plt.figure(figsize=(10, 5))
        plt.plot(self.uptime_df['Date'], self.uptime_df['Uptime'], marker='o')
        plt.title('Daily Uptime of Windows 11 Laptop')
        plt.xlabel('Date')
        plt.ylabel('Uptime (Hours)')
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.show()

    def export_to_csv(self):
        if hasattr(self, 'uptime_df'):
            filename = filedialog.asksaveasfilename(defaultextension='.csv',
                                                    filetypes=[('CSV Files', '*.csv')])
            if filename:
                self.uptime_df.to_csv(filename, index=False)
                messagebox.showinfo('Export Successful', f'Data exported to {filename}')
        else:
            messagebox.showwarning('No Data', 'Please track the uptime first.')

    def prefill_dates(self):
        # Get the current date and the date from a month before
        current_date = datetime.now()
        one_month_ago = current_date - timedelta(days=30)

        # Format dates as mm/dd/yyyy
        self.start_date_entry.insert(0, one_month_ago.strftime('%m/%d/%Y'))
        self.end_date_entry.insert(0, current_date.strftime('%m/%d/%Y'))
    def run(self):
        self.root.mainloop()

if __name__ == '__main__':
    root = tk.Tk()
    app = UptimeTrackerApp(root)
    app.run()
