"""Link the latest saved report and document versions in the ABC Ltd working papers."""
import argparse

from report_filing import link_existing_reports
from working_papers_link import WorkbookOpenError


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.parse_args()
    try:
        link_existing_reports()
    except WorkbookOpenError as error:
        parser.exit(3, f'{error}\n')
    except (OSError, ValueError) as error:
        parser.exit(1, f'Working-paper links could not be updated: {error}\n')


if __name__ == '__main__':
    main()
