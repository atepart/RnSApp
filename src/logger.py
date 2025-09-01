import logging


def configure_logger():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(name)s] [%(levelname)s] %(message)s",
        handlers=[logging.StreamHandler(), logging.FileHandler("debug.log")],
    )
