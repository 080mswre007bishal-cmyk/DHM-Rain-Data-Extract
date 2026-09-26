def extract_popup_rainfall(popup):
    """
    Complete operation on one clicked station popup:

        Leave Date unchanged
        -> Daily
        -> VIEW
        -> read table
        -> filter yesterday
    """

    target_date = yesterday_iso()

    print(
        f"    Target AD date: {target_date}"
    )

    # ========================================================
    # DO NOT CHANGE THE DATE
    # ========================================================

    print(
        "    Leaving popup Date unchanged."
    )

    # ========================================================
    # SET PERIOD = DAILY
    # ========================================================

    daily_ok = set_popup_period_daily(
        popup
    )

    if not daily_ok:

        # Save the popup HTML/text so we can inspect
        # exactly what DHM generated.
        if DEBUG:

            try:

                DEBUG_DIR.mkdir(
                    parents=True,
                    exist_ok=True,
                )

                (
                    DEBUG_DIR
                    / "period_dropdown_failed.html"
                ).write_text(
                    popup.evaluate(
                        "e => e.outerHTML"
                    ),
                    encoding="utf-8",
                )

                (
                    DEBUG_DIR
                    / "period_dropdown_failed.txt"
                ).write_text(
                    popup.inner_text(),
                    encoding="utf-8",
                )

            except Exception:
                pass

        raise RuntimeError(
            "Period dropdown was found, "
            "but could not be changed to Daily."
        )

    # ========================================================
    # CLICK VIEW
    # ========================================================

    if not click_popup_view(
        popup
    ):

        raise RuntimeError(
            "VIEW button was not found "
            "inside the station popup."
        )

    # ========================================================
    # WAIT FOR TABLE
    # ========================================================

    for _ in range(30):

        rows = (
            read_rainfall_table_from_popup(
                popup
            )
        )

        if rows:

            # Keep ONLY yesterday.
            yesterday_rows = [
                row
                for row in rows
                if row[0] == target_date
            ]

            if yesterday_rows:
                return yesterday_rows

        time.sleep(0.5)

    raise RuntimeError(
        "Daily rainfall table loaded, "
        f"but no row was found for "
        f"yesterday ({target_date})."
    )