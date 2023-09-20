from excel_processor import ExcelProcessor

INPUT_FOLDER_PATH = "./data_files"
OUTPUT_CSV_PATH = "./output_files/materials.csv"
OUTPUT_READ_PATH = "./output_files/read.txt"


def main():
    processor = ExcelProcessor(INPUT_FOLDER_PATH, OUTPUT_CSV_PATH, OUTPUT_READ_PATH)
    processor.process_excel_files()


if __name__ == "__main__":
    main()
