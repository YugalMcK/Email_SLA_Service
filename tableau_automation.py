"""
Automate McKesson Tableau export for Ops Only IB Summary Report.

Uses Playwright with a visible browser to complete SSO/Okta login, navigate the
Tableau portal, download a Crosstab Excel export, and save it to excel_reports/.
"""

import logging
import os
import time
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from config import (
    AUTH_STATE_PATH,
    DATA_REFRESH_TIMEOUT_MS,
    DOWNLOAD_TIMEOUT_MS,
    EXCEL_REPORT_FOLDER,
    LOGIN_TIMEOUT_MS,
    NoDataError,
    PAGE_LOAD_TIMEOUT_MS,
    get_browser_channels,
    get_browser_executable,
    get_date_range,
    get_report_url,
    get_tableau_password,
    get_tableau_username,
    timestamp_label,
)

logger = logging.getLogger("GR_Tableau.tableau")

TABLEAU_READY_SELECTORS = (
    # Landing-page indicators (when navigating via the portal search page):
    'input.QueryBox[aria-label="Keyword Search"]',
    'input[type="text"].QueryBox[aria-label="Keyword Search"]',
    '[data-tb-test-id="site-navigation"]',
    '[class*="SiteNav"]',
    'button[aria-label="Search"]',
    # Report/viz-page indicators (when navigating directly to a report URL):
    '[data-tb-test-id="viz-viewer-toolbar-button-download"]',
    'button#download',
    '.tab-vizHeaderWrapper',
    'div.tabToolbar',
    '[data-tb-test-id="view-toolbar"]',
)

VIZ_LOADING_SELECTORS = (
    '[data-tb-test-id="Spinner"]',
    ".tab-refresh-overlay",
    '[class*="loading"]',
    '[class*="Loading"]',
    '[class*="spinner"]',
    '[class*="Spinner"]',
)

REFRESH_DATA_SELECTORS = (
    'button:has(svg[data-tb-test-id="tb-icons-DatasourceCanRefreshIcon"])',
    'button:has([data-tb-test-id="tb-icons-DatasourceCanRefreshIcon"])',
    'svg[data-tb-test-id="tb-icons-DatasourceCanRefreshIcon"]',
    '[data-tb-test-id="tb-icons-DatasourceCanRefreshIcon"]',
)

DOWNLOAD_BUTTON_SELECTORS = (
    'button#download[data-tb-test-id="viz-viewer-toolbar-button-download"]',
    '[data-tb-test-id="viz-viewer-toolbar-button-download"]',
    'button#download',
    'button:has([data-tb-test-id="tb-icons-DownloadBaseIcon"])',
    'svg[data-tb-test-id="tb-icons-DownloadBaseIcon"]',
    '[data-tb-test-id="tb-icons-DownloadBaseIcon"]',
)

CROSSTAB_OPTION_SELECTORS = (
    'span.frvoegc:has-text("Crosstab")',
    'span:has-text("Crosstab")',
    '[role="menuitem"]:has-text("Crosstab")',
)

# Tableau shows this message when the current date/view has no data to export.
NO_DATA_TEXT = "No sheets to select. Try a different view."
NO_DATA_SELECTORS = (
    f'label:has-text("{NO_DATA_TEXT}")',
    f':text("{NO_DATA_TEXT}")',
    'label:has-text("No sheets to select")',
    ':text("No sheets to select")',
)

EXPORT_CROSSTAB_BUTTON_SELECTORS = (
    '[data-tb-test-id="export-crosstab-export-Button"]',
    'button[data-tb-test-id="export-crosstab-export-Button"]',
)

OKTA_VERIFICATION_SELECTORS = (
    "#okta-sign-in",
    "#sign-in-widget",
    '[data-se="o-form"]',
    '[data-se="factor-beacon"]',
    ".authenticator-row",
    ".authenticator-button",
    ".okta-form-subtitle",
    'div:has-text("Verify it\'s you")',
    'div:has-text("Verify your identity")',
    'div:has-text("Okta Verify")',
    'button:has-text("Send push")',
    'button:has-text("Receive a code")',
    'button:has-text("Get a call")',
    'input[name="answer"]',
    "#idDiv_SAOTCS_Proofs",
    'div:has-text("Approve sign in request")',
    'div:has-text("Enter code")',
    '[data-se="mfa-otp-challenge"]',
    '[data-se="mfa-verify-passcode"]',
)

LOGIN_FIELD_SELECTORS = (
    "#okta-signin-username",
    "#identifier",
    "input[name='username']",
    "input[name='identifier']",
    "input[type='email']",
)

PASSWORD_FIELD_SELECTORS = (
    "#okta-signin-password",
    "input[name='password']",
    "input[type='password']",
)


def _launch_browser(playwright):
    # If a specific (e.g. portable) Chrome executable is configured, use it directly.
    executable = get_browser_executable()
    if executable:
        logger.info("Launching browser from executable: %s", executable)
        return playwright.chromium.launch(executable_path=executable, headless=False)

    last_error = None
    for channel in get_browser_channels():
        label = channel or "chromium"
        try:
            logger.info("Launching browser: %s", label)
            if channel:
                browser = playwright.chromium.launch(channel=channel, headless=False)
            else:
                browser = playwright.chromium.launch(headless=False)
            logger.info("Browser started (%s).", label)
            return browser
        except Exception as exc:
            last_error = exc
            logger.warning("Could not launch %s: %s", label, exc)
    raise RuntimeError(f"Unable to launch a browser. Last error: {last_error}")


def _iter_contexts(page):
    """Yield the page and every frame (Okta often renders inside an iframe)."""
    yield page
    for frame in page.frames:
        if frame is not page.main_frame:
            yield frame


def _is_visible_in_contexts(page, selectors, *, timeout_ms: int = 500) -> bool:
    for context in _iter_contexts(page):
        for selector in selectors:
            try:
                if context.locator(selector).first.is_visible(timeout=timeout_ms):
                    return True
            except PlaywrightTimeoutError:
                continue
            except Exception:
                continue
    return False


def _is_sso_url(url: str) -> bool:
    lowered = url.lower()
    return any(
        marker in lowered
        for marker in (
            "okta.com",
            "microsoftonline.com",
            "login.microsoftonline",
            "/login",
            "/sso",
            "/signin",
        )
    )


def _is_login_page(page) -> bool:
    if _is_sso_url(page.url):
        return True
    return (
        _is_visible_in_contexts(page, LOGIN_FIELD_SELECTORS, timeout_ms=500)
        or _is_visible_in_contexts(page, PASSWORD_FIELD_SELECTORS, timeout_ms=500)
    )


def _is_password_field_visible(page) -> bool:
    return _is_visible_in_contexts(page, PASSWORD_FIELD_SELECTORS, timeout_ms=500)


def _is_okta_verification_pending(page) -> bool:
    """Return True while Okta/MFA factor selection or approval is still showing."""
    if _is_sso_url(page.url):
        return True
    return _is_visible_in_contexts(page, OKTA_VERIFICATION_SELECTORS, timeout_ms=500)


def _is_tableau_ready(page) -> bool:
    """Return True only when Tableau content is loaded and SSO is fully complete."""
    if "viz.mckesson.com" not in page.url.lower():
        return False
    if _is_login_page(page):
        return False
    if _is_okta_verification_pending(page):
        return False
    return _is_visible_in_contexts(page, TABLEAU_READY_SELECTORS, timeout_ms=2000)


def _active_page_for_auth(context, preferred_page):
    """Prefer the tab that currently shows Okta; otherwise use the landing page."""
    for candidate in reversed(context.pages):
        if _is_sso_url(candidate.url) or _is_okta_verification_pending(candidate):
            return candidate
    return preferred_page


def _save_auth_state(context) -> None:
    context.storage_state(path=str(AUTH_STATE_PATH))
    logger.info("Saved Tableau session to '%s'", AUTH_STATE_PATH)


def _new_browser_context(browser):
    if os.environ.get("TABLEAU_FORCE_LOGIN", "").strip().lower() in ("1", "true", "yes"):
        if AUTH_STATE_PATH.is_file():
            AUTH_STATE_PATH.unlink()
            logger.info("TABLEAU_FORCE_LOGIN set; removed saved session.")

    context_kwargs = {"accept_downloads": True}
    if AUTH_STATE_PATH.is_file():
        logger.info("Loading saved Tableau session from '%s'", AUTH_STATE_PATH)
        context_kwargs["storage_state"] = str(AUTH_STATE_PATH)

    context = browser.new_context(**context_kwargs)
    context.set_default_timeout(PAGE_LOAD_TIMEOUT_MS)
    return context


def _click_login_submit(page) -> bool:
    submit_selectors = (
        "#okta-signin-submit",
        'button:has-text("Next")',
        'button:has-text("Verify")',
        'button:has-text("Sign in")',
        'button:has-text("Log in")',
        'button:has-text("Continue")',
        'input[type="submit"]',
        'button[type="submit"]',
    )
    for selector in submit_selectors:
        submit = page.locator(selector).first
        try:
            if submit.is_visible(timeout=500):
                submit.click()
                page.wait_for_load_state("domcontentloaded", timeout=PAGE_LOAD_TIMEOUT_MS)
                return True
        except PlaywrightTimeoutError:
            continue
    return False


def _fill_first_empty_field(page, selectors, value: str, label: str) -> bool:
    for selector in selectors:
        field = page.locator(selector).first
        try:
            if not field.is_visible(timeout=500):
                continue
            if field.input_value().strip():
                continue
            logger.info("Auto-filling %s on '%s'.", label, selector)
            field.fill(value)
            return True
        except PlaywrightTimeoutError:
            continue
        except Exception:
            continue
    return False


def _try_auto_fill_login(page, username: str, password: str = "") -> bool:
    """Fill username always. Password is optional (service account only)."""
    if not username:
        return False

    username_filled = _fill_first_empty_field(
        page, LOGIN_FIELD_SELECTORS, username, "username"
    )
    password_filled = False
    if password:
        password_filled = _fill_first_empty_field(
            page, PASSWORD_FIELD_SELECTORS, password, "password"
        )

    if password_filled:
        _click_login_submit(page)
        logger.info("Service account login submitted with username and password.")
        return True

    if username_filled:
        _click_login_submit(page)
        if password:
            logger.info(
                "Username submitted for service account; waiting for password step."
            )
        else:
            logger.info(
                "Username submitted; complete Okta verification in the browser."
            )
        return True

    return False


def _wait_for_tableau_session(page, context, username: str = "", password: str = "") -> None:
    deadline = time.time() + (LOGIN_TIMEOUT_MS / 1000)
    verification_notice_logged = False

    while time.time() < deadline:
        active_page = _active_page_for_auth(context, page)

        for candidate in (page, active_page, *context.pages):
            if _is_tableau_ready(candidate):
                candidate.bring_to_front()
                _save_auth_state(context)
                logger.info("Tableau session is active.")
                return

        if _is_okta_verification_pending(active_page):
            if not verification_notice_logged:
                logger.info(
                    "Okta verification is required. Select your verification option "
                    "(Push, Okta Verify, SMS, etc.) and complete it in the browser. "
                    "Waiting up to %s minutes.",
                    LOGIN_TIMEOUT_MS // 60000,
                )
                verification_notice_logged = True
        elif _is_login_page(active_page):
            _try_auto_fill_login(active_page, username, password)
        elif password and _is_password_field_visible(active_page):
            _try_auto_fill_login(active_page, username, password)

        active_page.wait_for_timeout(1000)

    raise TimeoutError(
        "Tableau/Okta login was not completed within 5 minutes. "
        "Complete Okta verification in the browser window and rerun if needed."
    )


def ensure_tableau_session(page, context, landing_url: str) -> None:
    username = get_tableau_username()
    password = get_tableau_password()

    page.goto(landing_url, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT_MS)
    page.wait_for_timeout(3000)

    if _is_tableau_ready(page):
        logger.info("Tableau session is active.")
        _save_auth_state(context)
        return

    logger.info("Starting Tableau/Okta login.")
    logger.info("Using username: %s", username)
    if password:
        logger.info("Password is configured; using service account login flow.")
    else:
        logger.info(
            "No password configured; using Okta SSO flow after username submission."
        )

    _try_auto_fill_login(_active_page_for_auth(context, page), username, password)

    if not password:
        logger.info(
            "Complete Okta verification in the browser when prompted "
            "(Push, Okta Verify, SMS, etc.)."
        )

    _wait_for_tableau_session(page, context, username, password)


def _find_locator_in_contexts(page, selectors, *, timeout_ms: int | None = None):
    """Find the first visible locator in the page or any Tableau iframe."""
    timeout_ms = timeout_ms or PAGE_LOAD_TIMEOUT_MS
    deadline = time.time() + (timeout_ms / 1000)

    while time.time() < deadline:
        for context in _iter_contexts(page):
            for selector in selectors:
                locator = context.locator(selector).first
                try:
                    if locator.is_visible(timeout=500):
                        return locator
                except PlaywrightTimeoutError:
                    continue
                except Exception:
                    continue
        page.wait_for_timeout(1000)

    return None


def _settle(page, *, extra_ms: int = 2000) -> None:
    """Wait for the DOM to be ready, then pause briefly.

    NOTE: Tableau vizzes keep network connections open (websockets/polling), so
    'networkidle' NEVER fires and must not be used - it always times out.
    """
    try:
        page.wait_for_load_state("domcontentloaded", timeout=PAGE_LOAD_TIMEOUT_MS)
    except Exception as exc:  # noqa: BLE001 - non-fatal, keep going
        logger.warning("wait_for_load_state('domcontentloaded') failed: %s", exc)
    page.wait_for_timeout(extra_ms)


def _is_viz_loading(page) -> bool:
    return _is_visible_in_contexts(page, VIZ_LOADING_SELECTORS, timeout_ms=300)


def _wait_for_data_load(page, *, label: str = "data") -> None:
    """Wait until Tableau finishes loading after a parameter change."""
    logger.info("Waiting for %s to finish loading.", label)
    page.wait_for_timeout(1500)

    deadline = time.time() + (DATA_REFRESH_TIMEOUT_MS / 1000)
    stable_checks = 0
    while time.time() < deadline:
        if _is_viz_loading(page):
            stable_checks = 0
        else:
            stable_checks += 1
            if stable_checks >= 2:
                logger.info("%s load completed.", label)
                page.wait_for_timeout(1000)
                return
        page.wait_for_timeout(1000)

    logger.warning("Timed out waiting for %s to load; continuing.", label)


def _fill_date_field(page, aria_label: str, value: str) -> None:
    selectors = (
        f'input.QueryBox[aria-label="{aria_label}"]',
        f'input[type="text"].QueryBox[aria-label="{aria_label}"]',
    )
    field = _find_locator_in_contexts(page, selectors)
    if field is None:
        raise RuntimeError(f'Could not find QueryBox input with aria-label="{aria_label}".')

    field.wait_for(state="visible", timeout=PAGE_LOAD_TIMEOUT_MS)
    field.scroll_into_view_if_needed()
    field.click()
    field.fill("")
    field.fill(value)
    field.press("Tab")
    logger.info("Entered date '%s' in '%s' field.", value, aria_label)


def _set_date_range(page, from_date: str, to_date: str) -> None:
    logger.info("Setting date range From=%s To=%s", from_date, to_date)
    _fill_date_field(page, "From", from_date)
    _wait_for_data_load(page, label="From date filter")
    _fill_date_field(page, "To", to_date)
    _wait_for_data_load(page, label="To date filter")


def _click_locator_or_parent_button(locator) -> None:
    locator.wait_for(state="visible", timeout=PAGE_LOAD_TIMEOUT_MS)
    locator.scroll_into_view_if_needed()
    parent_button = locator.locator("xpath=ancestor::button[1]")
    try:
        if parent_button.count() > 0 and parent_button.is_visible(timeout=500):
            parent_button.click(timeout=PAGE_LOAD_TIMEOUT_MS)
            return
    except PlaywrightTimeoutError:
        pass
    locator.click(timeout=PAGE_LOAD_TIMEOUT_MS)


def _click_refresh_data(page) -> None:
    logger.info("Clicking refresh data control (DatasourceCanRefreshIcon).")
    refresh = _find_locator_in_contexts(page, REFRESH_DATA_SELECTORS)
    if refresh is None:
        raise RuntimeError(
            'Could not find refresh data control '
            '[data-tb-test-id="tb-icons-DatasourceCanRefreshIcon"].'
        )

    _click_locator_or_parent_button(refresh)
    logger.info("Refresh data control clicked.")
    _wait_for_data_load(page, label="manual refresh")


def _check_no_data(page) -> None:
    """Raise NoDataError if Tableau reports there are no sheets/data to export."""
    if _is_visible_in_contexts(page, NO_DATA_SELECTORS, timeout_ms=1000):
        raise NoDataError(
            f'Tableau reported "{NO_DATA_TEXT}" - there is no data for the '
            f"selected date. Terminating without downloading."
        )


def _download_crosstab(page, download_dir: Path) -> Path:
    logger.info("Starting Crosstab download.")

    download_button = _find_locator_in_contexts(page, DOWNLOAD_BUTTON_SELECTORS)
    if download_button is None:
        raise RuntimeError(
            'Could not find download button '
            '[data-tb-test-id="viz-viewer-toolbar-button-download"].'
        )

    _click_locator_or_parent_button(download_button)
    logger.info("Download toolbar button clicked.")
    page.wait_for_timeout(1000)

    crosstab = _find_locator_in_contexts(page, CROSSTAB_OPTION_SELECTORS)
    if crosstab is None:
        raise RuntimeError('Could not find "Crosstab" option in the download menu.')
    crosstab.click(timeout=PAGE_LOAD_TIMEOUT_MS)
    logger.info("Crosstab export option selected.")
    page.wait_for_timeout(1000)

    # The Crosstab dialog shows "No sheets to select. Try a different view."
    # when the selected date has no data. Detect it and stop gracefully.
    _check_no_data(page)

    export_button = _find_locator_in_contexts(page, EXPORT_CROSSTAB_BUTTON_SELECTORS)
    if export_button is None:
        raise RuntimeError(
            'Could not find Crosstab export Download button '
            '[data-tb-test-id="export-crosstab-export-Button"].'
        )
    export_button.wait_for(state="visible", timeout=PAGE_LOAD_TIMEOUT_MS)

    download_dir.mkdir(parents=True, exist_ok=True)
    target_path = download_dir / f"Ops_Only_IB_Summary_{timestamp_label()}.xlsx"

    with page.expect_download(timeout=DOWNLOAD_TIMEOUT_MS) as download_info:
        export_button.click()

    download = download_info.value
    download.save_as(str(target_path))
    logger.info("Download saved to '%s'", target_path)
    return target_path


def download_tableau_report(
    *,
    landing_url: str | None = None,
    download_dir: Path | None = None,
) -> Path:
    landing_url = landing_url or get_report_url()
    download_dir = download_dir or EXCEL_REPORT_FOLDER
    from_date, to_date = get_date_range()

    with sync_playwright() as playwright:
        browser = _launch_browser(playwright)
        context = _new_browser_context(browser)
        page = context.new_page()

        try:
            # The report_url points DIRECTLY at the report ("second tab"), so we
            # navigate straight to it and log in if required. The old first-tab
            # steps (keyword search + clicking the landing report) are skipped.
            ensure_tableau_session(page, context, landing_url)

            workbook_page = page
            workbook_page.bring_to_front()
            _settle(workbook_page, extra_ms=2000)
            logger.info("Report page loaded: %s", workbook_page.url)

            _set_date_range(workbook_page, from_date, to_date)
            _click_refresh_data(workbook_page)
            downloaded_file = _download_crosstab(workbook_page, download_dir)
            _save_auth_state(context)
            return downloaded_file
        finally:
            context.close()
            browser.close()
