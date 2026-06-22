#!/usr/bin/env python3
"""
CONSAR Data Pipeline Orchestrator
================================

Orchestrates the complete CONSAR data collection and processing pipeline:
1. Downloads data using main.py (CONSARDownloader)
2. Processes and maps data using map.py (CleanUniversalMexicanPensionMapper)
3. Generates final structured output

Features:
✅ Complete automation from download to final output
✅ Error handling and recovery
✅ Progress tracking and detailed logging
✅ File cleanup and organization
✅ Configurable settings
✅ Pipeline status reporting

Dependencies: All dependencies from main.py and map.py
"""

import os
import sys
import time
import shutil
from datetime import datetime
from pathlib import Path

# Import the downloader and mapper classes
from main import CONSARDownloader
from map import CleanUniversalMexicanPensionMapper


class CONSARPipelineOrchestrator:
    """
    Orchestrates the complete CONSAR data pipeline from download to processing.
    """

    def __init__(self, base_directory=None, cleanup_downloads=True):
        """
        Initialize the pipeline orchestrator.

        Args:
            base_directory (str): Base directory for all operations
            cleanup_downloads (bool): Whether to clean up raw download files after processing
        """
        self.base_directory = Path(base_directory) if base_directory else Path.cwd()
        self.cleanup_downloads = cleanup_downloads

        # Create directory structure
        self.downloads_dir = self.base_directory / "downloads"
        self.processed_dir = self.base_directory / "output"
        self.logs_dir = self.base_directory / "logs"

        # Create directories
        for directory in [self.downloads_dir, self.processed_dir, self.logs_dir]:
            directory.mkdir(exist_ok=True)

        # Initialize components
        self.downloader = CONSARDownloader(str(self.downloads_dir))
        self.mapper = CleanUniversalMexicanPensionMapper()

        # Pipeline results
        self.pipeline_results = {
            'start_time': None,
            'end_time': None,
            'downloads': {},
            'processing_results': None,
            'final_output_file': None,
            'errors': [],
            'warnings': []
        }

    def log(self, message, level="INFO"):
        """Enhanced logging with timestamps and levels."""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        levels = {"INFO": "ℹ️", "SUCCESS": "✅", "WARNING": "⚠️", "ERROR": "❌", "PROGRESS": "🔄"}
        symbol = levels.get(level, "ℹ️")

        log_message = f"[{timestamp}] {symbol} {message}"
        print(log_message)

        # Write to log file
        log_file = self.logs_dir / f"pipeline_{datetime.now().strftime('%Y%m%d')}.log"
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(log_message + '\n')

    def download_all_data(self):
        """Execute all three download steps."""
        self.log("Starting CONSAR data download process...", "PROGRESS")

        download_results = {}

        # Step 1: Investment Data (monthly chart)
        self.log("Step 1: Downloading Investment Data...", "PROGRESS")
        try:
            investment_file = self.downloader.download_monthly_chart()
            if investment_file:
                file_size = os.path.getsize(investment_file)
                download_results['investment'] = investment_file
                self.log(f"Investment data downloaded: {os.path.basename(investment_file)} ({file_size:,} bytes)", "SUCCESS")
            else:
                download_results['investment'] = None
                self.log("Investment data download failed", "ERROR")
                self.pipeline_results['errors'].append("Investment data download failed")
        except Exception as e:
            download_results['investment'] = None
            self.log(f"Investment data download error: {e}", "ERROR")
            self.pipeline_results['errors'].append(f"Investment download error: {e}")

        # Step 2: RCV Flow Data
        self.log("Step 2: Downloading RCV Flow Data...", "PROGRESS")
        try:
            flow_file = self.downloader.download_rcv_flow_data()
            if flow_file:
                file_size = os.path.getsize(flow_file)
                download_results['rcv_flow'] = flow_file
                self.log(f"RCV Flow data downloaded: {os.path.basename(flow_file)} ({file_size:,} bytes)", "SUCCESS")
            else:
                download_results['rcv_flow'] = None
                self.log("RCV Flow data download failed", "ERROR")
                self.pipeline_results['errors'].append("RCV Flow data download failed")
        except Exception as e:
            download_results['rcv_flow'] = None
            self.log(f"RCV Flow data download error: {e}", "ERROR")
            self.pipeline_results['errors'].append(f"RCV Flow download error: {e}")

        # Step 3: Withdrawal Data
        self.log("Step 3: Downloading Withdrawal Data...", "PROGRESS")
        try:
            withdrawal_file = self.downloader.download_withdrawal_data()
            if withdrawal_file:
                file_size = os.path.getsize(withdrawal_file)
                download_results['withdrawal'] = withdrawal_file
                self.log(f"Withdrawal data downloaded: {os.path.basename(withdrawal_file)} ({file_size:,} bytes)", "SUCCESS")
            else:
                download_results['withdrawal'] = None
                self.log("Withdrawal data download failed", "ERROR")
                self.pipeline_results['errors'].append("Withdrawal data download failed")
        except Exception as e:
            download_results['withdrawal'] = None
            self.log(f"Withdrawal data download error: {e}", "ERROR")
            self.pipeline_results['errors'].append(f"Withdrawal download error: {e}")

        # Summary
        successful_downloads = sum(1 for result in download_results.values() if result is not None)
        self.log(f"Download phase completed: {successful_downloads}/3 files downloaded",
                "SUCCESS" if successful_downloads == 3 else "WARNING")

        self.pipeline_results['downloads'] = download_results
        return download_results

    def process_data(self, download_results):
        """Process the downloaded data using the mapper."""
        self.log("Starting data processing and mapping...", "PROGRESS")

        if not download_results.get('investment'):
            self.log("Cannot process: No investment file available", "ERROR")
            return None

        try:
            # Collect flow files
            flow_files = []
            for key in ['rcv_flow', 'withdrawal']:
                if download_results.get(key):
                    flow_files.append(download_results[key])

            # Process with the mapper
            self.log(f"Processing files: 1 investment + {len(flow_files)} flow files", "INFO")

            processing_results = self.mapper.process_files(
                investment_file=download_results['investment'],
                flow_files=flow_files,
                directory=str(self.downloads_dir)
            )

            # Log processing results
            self.log(f"Data processing completed:", "SUCCESS")
            self.log(f"  • Investment categories: {processing_results['investment_categories']}", "INFO")
            self.log(f"  • Flow categories: {processing_results['flow_categories']}", "INFO")
            self.log(f"  • Total mapped: {processing_results['total_categories']}/50", "INFO")
            self.log(f"  • Data date: {processing_results['date']}", "INFO")

            self.pipeline_results['processing_results'] = processing_results
            return processing_results

        except Exception as e:
            self.log(f"Data processing error: {e}", "ERROR")
            self.pipeline_results['errors'].append(f"Data processing error: {e}")
            return None

    def export_final_output(self, processing_results):
        """Export final processed output to the processed directory."""
        if not processing_results or processing_results['consar_data'].empty:
            self.log("Cannot export: No processed data available", "ERROR")
            return None

        try:
            # Generate output filename with timestamp
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            data_date = processing_results.get('date', 'unknown').replace('-', '')
            output_filename = f"CONSAR_Pipeline_Output_{data_date}_{timestamp}.xlsx"
            output_path = self.processed_dir / output_filename

            # Export using mapper's export function
            self.mapper.export_results(processing_results, str(output_path))

            self.log(f"Final output exported: {output_filename}", "SUCCESS")
            self.pipeline_results['final_output_file'] = str(output_path)
            return str(output_path)

        except Exception as e:
            self.log(f"Export error: {e}", "ERROR")
            self.pipeline_results['errors'].append(f"Export error: {e}")
            return None

    def cleanup_download_files(self):
        """Clean up raw download files after successful processing."""
        if not self.cleanup_downloads:
            self.log("Cleanup skipped (disabled)", "INFO")
            return

        try:
            # Only cleanup if we have a successful final output
            if not self.pipeline_results.get('final_output_file'):
                self.log("Cleanup skipped: No successful final output", "WARNING")
                return

            # Move download files to a backup folder
            backup_dir = self.downloads_dir / f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            backup_dir.mkdir(exist_ok=True)

            files_moved = 0
            for download_file in self.pipeline_results['downloads'].values():
                if download_file and os.path.exists(download_file):
                    try:
                        shutil.move(download_file, backup_dir / os.path.basename(download_file))
                        files_moved += 1
                    except Exception as e:
                        self.log(f"Could not move {download_file}: {e}", "WARNING")

            self.log(f"Cleanup completed: {files_moved} files moved to backup", "SUCCESS")

        except Exception as e:
            self.log(f"Cleanup error: {e}", "WARNING")

    def generate_pipeline_report(self):
        """Generate a comprehensive pipeline execution report."""
        try:
            # Calculate execution time
            if self.pipeline_results['start_time'] and self.pipeline_results['end_time']:
                execution_time = self.pipeline_results['end_time'] - self.pipeline_results['start_time']
                execution_minutes = execution_time.total_seconds() / 60
            else:
                execution_minutes = "Unknown"

            # Create report
            report_data = [
                ["CONSAR Pipeline Execution Report", ""],
                ["=" * 50, ""],
                ["Execution Date", datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                ["Execution Time (minutes)", f"{execution_minutes:.2f}" if isinstance(execution_minutes, float) else execution_minutes],
                ["Base Directory", str(self.base_directory)],
                ["", ""],
                ["DOWNLOAD RESULTS", ""],
                ["-" * 20, ""],
                ["Investment File", "✅ Success" if self.pipeline_results['downloads'].get('investment') else "❌ Failed"],
                ["RCV Flow File", "✅ Success" if self.pipeline_results['downloads'].get('rcv_flow') else "❌ Failed"],
                ["Withdrawal File", "✅ Success" if self.pipeline_results['downloads'].get('withdrawal') else "❌ Failed"],
                ["", ""],
                ["PROCESSING RESULTS", ""],
                ["-" * 20, ""],
            ]

            if self.pipeline_results['processing_results']:
                pr = self.pipeline_results['processing_results']
                report_data.extend([
                    ["Data Date", pr.get('date', 'Not extracted')],
                    ["Investment Categories", pr.get('investment_categories', 0)],
                    ["Flow Categories", pr.get('flow_categories', 0)],
                    ["Total Categories Mapped", f"{pr.get('total_categories', 0)}/50"],
                    ["Processing Status", "✅ Success"],
                ])
            else:
                report_data.append(["Processing Status", "❌ Failed"])

            report_data.extend([
                ["", ""],
                ["OUTPUT", ""],
                ["-" * 20, ""],
                ["Final Output File", self.pipeline_results.get('final_output_file', 'Not generated')],
                ["", ""],
                ["ISSUES", ""],
                ["-" * 20, ""],
                ["Errors", len(self.pipeline_results['errors'])],
                ["Warnings", len(self.pipeline_results['warnings'])],
            ])

            # Add errors and warnings
            for i, error in enumerate(self.pipeline_results['errors'], 1):
                report_data.append([f"Error {i}", error])

            for i, warning in enumerate(self.pipeline_results['warnings'], 1):
                report_data.append([f"Warning {i}", warning])

            # Save report
            report_filename = f"Pipeline_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            report_path = self.logs_dir / report_filename

            with open(report_path, 'w', encoding='utf-8') as f:
                for row in report_data:
                    f.write(f"{row[0]:<30} {row[1]}\n")

            self.log(f"Pipeline report generated: {report_filename}", "SUCCESS")
            return str(report_path)

        except Exception as e:
            self.log(f"Report generation error: {e}", "ERROR")
            return None

    def run_complete_pipeline(self):
        """Execute the complete CONSAR data pipeline."""
        self.pipeline_results['start_time'] = datetime.now()

        self.log("=" * 80, "INFO")
        self.log("CONSAR DATA PIPELINE ORCHESTRATOR", "INFO")
        self.log("=" * 80, "INFO")
        self.log("🚀 Starting complete pipeline execution...", "PROGRESS")

        try:
            # Phase 1: Download data
            self.log("\n📥 PHASE 1: DATA DOWNLOAD", "INFO")
            download_results = self.download_all_data()

            if not any(download_results.values()):
                self.log("Pipeline terminated: No files downloaded", "ERROR")
                return False

            # Phase 2: Process data
            self.log("\n🔄 PHASE 2: DATA PROCESSING", "INFO")
            processing_results = self.process_data(download_results)

            if not processing_results:
                self.log("Pipeline terminated: Data processing failed", "ERROR")
                return False

            # Phase 3: Export final output
            self.log("\n📤 PHASE 3: FINAL EXPORT", "INFO")
            final_output = self.export_final_output(processing_results)

            if not final_output:
                self.log("Pipeline terminated: Export failed", "ERROR")
                return False

            # Phase 4: Cleanup (optional)
            self.log("\n🧹 PHASE 4: CLEANUP", "INFO")
            self.cleanup_download_files()

            # Pipeline success
            self.pipeline_results['end_time'] = datetime.now()

            self.log("\n🎉 PIPELINE COMPLETED SUCCESSFULLY!", "SUCCESS")
            self.log(f"📁 Final output: {os.path.basename(final_output)}", "SUCCESS")
            self.log(f"📊 Data mapped: {processing_results['total_categories']}/50 categories", "SUCCESS")
            self.log(f"📅 Data date: {processing_results['date']}", "SUCCESS")

            # Generate report
            self.log("\n📋 GENERATING REPORT", "INFO")
            report_file = self.generate_pipeline_report()

            return True

        except Exception as e:
            self.pipeline_results['end_time'] = datetime.now()
            self.log(f"Pipeline failed with unexpected error: {e}", "ERROR")
            self.pipeline_results['errors'].append(f"Unexpected pipeline error: {e}")
            return False

        finally:
            # Always generate a report
            if not self.pipeline_results.get('end_time'):
                self.pipeline_results['end_time'] = datetime.now()
            self.generate_pipeline_report()


def main():
    """Main function - Execute CONSAR pipeline with command line options."""
    print("🌍 CONSAR Data Pipeline Orchestrator")
    print("=" * 50)

    # Configuration
    base_directory = "."
    cleanup_downloads = True

    # Parse command line arguments
    if len(sys.argv) > 1:
        base_directory = sys.argv[1]

    if len(sys.argv) > 2 and sys.argv[2].lower() == 'no-cleanup':
        cleanup_downloads = False

    # Initialize and run pipeline
    orchestrator = CONSARPipelineOrchestrator(
        base_directory=base_directory,
        cleanup_downloads=cleanup_downloads
    )

    success = orchestrator.run_complete_pipeline()

    if success:
        print(f"\n✨ Pipeline completed successfully!")
        print(f"📁 Check the 'processed' folder for final output")
        print(f"📋 Check the 'logs' folder for detailed reports")
    else:
        print(f"\n❌ Pipeline failed. Check logs for details.")
        sys.exit(1)


if __name__ == "__main__":
    main()