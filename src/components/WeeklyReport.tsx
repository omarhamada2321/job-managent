import { useState, useEffect } from 'react';
import { supabase } from '../lib/supabase';
import { useAuth } from '../contexts/AuthContext';
import { ArrowLeft, Copy, Download, CheckCircle2, Circle, FileText, FileJson } from 'lucide-react';
import { generateTextReport, generateWordReport, generatePdfReport, ReportData } from '../lib/reportGenerator';

interface Task {
  id: string;
  title: string;
  completed: boolean;
  date: string;
}

interface TasksByDate {
  [date: string]: Task[];
}

interface WeeklyReportProps {
  onNavigate: (view: 'daily' | 'weekly') => void;
}

export function WeeklyReport({ onNavigate }: WeeklyReportProps) {
  const [tasks, setTasks] = useState<TasksByDate>({});
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [loading, setLoading] = useState(false);
  const [copied, setCopied] = useState(false);
  const [exporting, setExporting] = useState(false);
  const { signOut } = useAuth();

  useEffect(() => {
    const today = new Date();
    const weekAgo = new Date(today);
    weekAgo.setDate(weekAgo.getDate() - 6);

    setStartDate(weekAgo.toISOString().split('T')[0]);
    setEndDate(today.toISOString().split('T')[0]);
  }, []);

  useEffect(() => {
    if (startDate && endDate) {
      fetchWeeklyTasks();
    }
  }, [startDate, endDate]);

  const fetchWeeklyTasks = async () => {
    setLoading(true);
    const { data, error } = await supabase
      .from('tasks')
      .select('*')
      .gte('date', startDate)
      .lte('date', endDate)
      .order('date', { ascending: true })
      .order('created_at', { ascending: true });

    if (error) {
      console.error('Error fetching tasks:', error);
    } else {
      const grouped: TasksByDate = {};
      (data || []).forEach((task) => {
        if (!grouped[task.date]) {
          grouped[task.date] = [];
        }
        grouped[task.date].push(task);
      });
      setTasks(grouped);
    }
    setLoading(false);
  };

  const formatDate = (dateStr: string) => {
    const date = new Date(dateStr + 'T00:00:00');
    return date.toLocaleDateString('en-US', {
      weekday: 'long',
      year: 'numeric',
      month: 'long',
      day: 'numeric',
    });
  };

  const getReportData = (): ReportData => {
    const dates = Object.keys(tasks).sort();
    let totalTasks = 0;
    let totalCompleted = 0;

    dates.forEach((date) => {
      const dayTasks = tasks[date];
      const completed = dayTasks.filter((t) => t.completed).length;
      totalTasks += dayTasks.length;
      totalCompleted += completed;
    });

    return {
      startDate,
      endDate,
      tasks,
      totalTasks,
      totalCompleted,
      completionRate: totalTasks > 0 ? Math.round((totalCompleted / totalTasks) * 100) : 0,
    };
  };

  const copyReport = async () => {
    const reportData = getReportData();
    const reportText = generateTextReport(reportData);
    await navigator.clipboard.writeText(reportText);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const downloadTextReport = () => {
    const reportData = getReportData();
    const reportText = generateTextReport(reportData);
    const blob = new Blob([reportText], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `weekly-report-${startDate}-to-${endDate}.txt`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const downloadWordReport = async () => {
    setExporting(true);
    try {
      const reportData = getReportData();
      await generateWordReport(reportData, `weekly-report-${startDate}-to-${endDate}.docx`);
    } catch (error) {
      console.error('Error generating Word report:', error);
    } finally {
      setExporting(false);
    }
  };

  const downloadPdfReport = () => {
    setExporting(true);
    try {
      const reportData = getReportData();
      generatePdfReport(reportData, `weekly-report-${startDate}-to-${endDate}.pdf`);
    } finally {
      setExporting(false);
    }
  };

  const dates = Object.keys(tasks).sort();
  const totalTasks = dates.reduce((sum, date) => sum + tasks[date].length, 0);
  const totalCompleted = dates.reduce(
    (sum, date) => sum + tasks[date].filter((t) => t.completed).length,
    0
  );

  return (
    <div className="min-h-screen bg-gradient-to-br from-green-50 via-white to-blue-50">
      <div className="max-w-5xl mx-auto p-6">
        <div className="bg-white rounded-2xl shadow-lg p-6">
          <div className="flex items-center justify-between mb-6">
            <button
              onClick={() => onNavigate('daily')}
              className="flex items-center gap-2 px-4 py-2 text-gray-600 hover:text-gray-800 transition"
            >
              <ArrowLeft className="w-5 h-5" />
              Back to Daily Tasks
            </button>
            <button
              onClick={() => signOut()}
              className="px-4 py-2 text-gray-600 hover:text-gray-800 transition"
            >
              Sign Out
            </button>
          </div>

          <h1 className="text-3xl font-bold text-gray-800 mb-6">Weekly Report</h1>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Start Date
              </label>
              <input
                type="date"
                value={startDate}
                onChange={(e) => setStartDate(e.target.value)}
                className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-transparent outline-none"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                End Date
              </label>
              <input
                type="date"
                value={endDate}
                onChange={(e) => setEndDate(e.target.value)}
                className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-transparent outline-none"
              />
            </div>
          </div>

          {totalTasks > 0 && (
            <div className="bg-green-50 border border-green-200 rounded-lg p-6 mb-6">
              <h3 className="font-semibold text-green-900 mb-3">Summary</h3>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-4 text-center">
                <div>
                  <p className="text-2xl font-bold text-green-700">{totalTasks}</p>
                  <p className="text-sm text-green-600">Total Tasks</p>
                </div>
                <div>
                  <p className="text-2xl font-bold text-green-700">{totalCompleted}</p>
                  <p className="text-sm text-green-600">Completed</p>
                </div>
                <div>
                  <p className="text-2xl font-bold text-orange-700">{totalTasks - totalCompleted}</p>
                  <p className="text-sm text-orange-600">Pending</p>
                </div>
                <div>
                  <p className="text-2xl font-bold text-blue-700">
                    {Math.round((totalCompleted / totalTasks) * 100)}%
                  </p>
                  <p className="text-sm text-blue-600">Completion Rate</p>
                </div>
              </div>
            </div>
          )}

          <div className="flex flex-wrap gap-3 mb-6">
            <button
              onClick={copyReport}
              className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg transition disabled:opacity-50 disabled:cursor-not-allowed"
              disabled={exporting}
            >
              <Copy className="w-4 h-4" />
              {copied ? 'Copied!' : 'Copy as Text'}
            </button>
            <button
              onClick={downloadTextReport}
              className="flex items-center gap-2 px-4 py-2 bg-slate-600 hover:bg-slate-700 text-white rounded-lg transition disabled:opacity-50 disabled:cursor-not-allowed"
              disabled={exporting}
            >
              <FileText className="w-4 h-4" />
              Download TXT
            </button>
            <button
              onClick={downloadWordReport}
              className="flex items-center gap-2 px-4 py-2 bg-blue-500 hover:bg-blue-600 text-white rounded-lg transition disabled:opacity-50 disabled:cursor-not-allowed"
              disabled={exporting}
            >
              <FileJson className="w-4 h-4" />
              {exporting ? 'Generating...' : 'Download DOCX'}
            </button>
            <button
              onClick={downloadPdfReport}
              className="flex items-center gap-2 px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded-lg transition disabled:opacity-50 disabled:cursor-not-allowed"
              disabled={exporting}
            >
              <Download className="w-4 h-4" />
              {exporting ? 'Generating...' : 'Download PDF'}
            </button>
          </div>

          {loading ? (
            <div className="text-center py-8 text-gray-500">Loading report...</div>
          ) : dates.length === 0 ? (
            <div className="text-center py-12 text-gray-500">
              <p className="text-lg">No tasks found for this period.</p>
            </div>
          ) : (
            <div className="space-y-6">
              {dates.map((date) => {
                const dayTasks = tasks[date];
                const completed = dayTasks.filter((t) => t.completed).length;

                return (
                  <div key={date} className="border border-gray-200 rounded-lg p-5">
                    <div className="flex items-center justify-between mb-4">
                      <h3 className="text-lg font-semibold text-gray-800">
                        {formatDate(date)}
                      </h3>
                      <span className="text-sm text-gray-600">
                        {completed}/{dayTasks.length} completed
                      </span>
                    </div>

                    <div className="space-y-2">
                      {dayTasks.map((task) => (
                        <div
                          key={task.id}
                          className="flex items-center gap-3 p-3 bg-gray-50 rounded"
                        >
                          {task.completed ? (
                            <CheckCircle2 className="w-5 h-5 text-green-600 flex-shrink-0" />
                          ) : (
                            <Circle className="w-5 h-5 text-gray-400 flex-shrink-0" />
                          )}
                          <span
                            className={
                              task.completed
                                ? 'line-through text-gray-500'
                                : 'text-gray-800'
                            }
                          >
                            {task.title}
                          </span>
                        </div>
                      ))}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
