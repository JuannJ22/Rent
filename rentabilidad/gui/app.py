from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

from nicegui import ui

from rentabilidad.config import bus, settings
from rentabilidad.app.dto import GenerarInformeRequest
from rentabilidad.app.use_cases.generar_informe_automatico import run as uc_auto
from rentabilidad.app.use_cases.generar_informe_manual import run as uc_manual


ui.add_head_html(
    """
<style>
  .q-card { border-radius: 1rem !important; }
  .q-field__control, .q-btn { border-radius: .75rem !important; }
  .q-btn { box-shadow: 0 1px 2px rgba(0,0,0,.06) !important; }
</style>
"""
)

card_classes = 'rounded-2xl shadow-sm border border-gray-200 bg-white'

state = SimpleNamespace(empty=None, log=None, last_update=None)


def agregar_log(msg: str) -> None:
    state.empty.add_class('hidden')
    state.log.remove_class('hidden')
    with state.log:
        ui.label(msg).classes('text-sm text-gray-700')


def touch_last_update() -> None:
    from datetime import datetime

    state.last_update.text = f"Última actualización: {datetime.now().strftime('%H:%M:%S')}"


bus.subscribe("log", agregar_log)
bus.subscribe("done", lambda m: (agregar_log(m), touch_last_update()))
bus.subscribe("error", lambda m: agregar_log(f"ERROR: {m}"))


with ui.header().classes("bg-[#1967d2] text-white"):
    with ui.row().classes("items-center justify-between w-full max-w-6xl mx-auto"):
        with ui.column():
            ui.label("Panel de rentabilidad").classes("text-2xl font-semibold leading-tight")
            ui.label("Control y automatización de procesos").classes("opacity-90 text-sm -mt-1")


with ui.column().classes("max-w-5xl mx-auto py-6 gap-6"):
    # --- Plantilla base (limpia, sin contenedor alto ni fondo) ---
    with ui.column().classes('gap-2 w-full'):
        with ui.row().classes('items-center gap-2'):
            ui.icon('folder_open').classes('text-gray-600')
            ui.label('Plantilla base').classes('font-medium')

        with ui.row().classes('items-center gap-2 w-full'):
            ruta_input = (
                ui.input(value=str(settings.ruta_plantilla))
                .props('readonly')
                .classes('flex-1 bg-gray-50 rounded-xl p-2 h-10 min-h-0 text-sm')
            )
            ui.button(
                'Copiar',
                on_click=lambda: ui.run_javascript(
                    f'navigator.clipboard.writeText("{settings.ruta_plantilla}")'
                ),
            )

            def abrir_carpeta():
                import os, sys

                p = Path(settings.ruta_plantilla).parent
                if sys.platform.startswith('win'):
                    os.startfile(p)  # type: ignore

            ui.button('Abrir carpeta', on_click=abrir_carpeta)

    with ui.card().classes(card_classes):
        ui.label('Informe automático').classes('font-medium')
        ui.label(
            'Genera automáticamente el informe del día anterior usando los archivos más recientes.'
        ).classes('text-sm text-gray-600')
        ui.button(
            'Generar informe automático',
            on_click=lambda: uc_auto(
                GenerarInformeRequest(ruta_plantilla=str(settings.ruta_plantilla)),
                bus,
            ),
        ).classes('mt-3 w-full')

    with ui.card().classes(card_classes):
        ui.label('Informe manual').classes('font-medium')
        ui.label('Genera un informe para una fecha específica.').classes('text-sm text-gray-600')
        fecha = ui.input(label='Fecha (YYYY-MM-DD)', value='2025-09-20').classes('w-48')
        ui.button(
            'Generar informe manual',
            on_click=lambda: uc_manual(
                GenerarInformeRequest(
                    ruta_plantilla=str(settings.ruta_plantilla),
                    fecha=fecha.value,
                ),
                bus,
            ),
        ).classes('mt-3 w-full')

    # --- Registro de Actividades ---
    with ui.card().classes(f'{card_classes}'):
        with ui.row().classes('items-center gap-2 px-4 pt-4'):
            ui.icon('activity').classes('text-violet-500')
            ui.label('Registro de Actividades').classes('font-medium')

            def limpiar_log():
                state.log.clear()
                state.log.add_class('hidden')
                state.empty.remove_class('hidden')

            ui.button('Limpiar', icon='delete', on_click=limpiar_log).props('flat').classes('ml-auto')

        with ui.element('div').classes('px-4 pb-4'):
            state.empty = ui.column().classes(
                'items-center justify-center h-40 w-full text-gray-400 bg-gray-50 rounded-xl'
            )
            with state.empty:
                ui.icon('inbox').classes('text-4xl')
                ui.label('El registro de actividades aparecerá aquí').classes('text-sm')

            state.log = ui.column().classes('hidden w-full gap-1 mt-3')

    # --- Footer ---
    with ui.row().classes('items-center justify-between text-xs text-gray-500'):
        with ui.row().classes('items-center gap-1'):
            ui.icon('check_circle').classes('text-emerald-500')
            ui.label('Sistema listo')
        from datetime import datetime

        state.last_update = ui.label(f'Última actualización: {datetime.now().strftime("%H:%M:%S")}')


def main() -> None:  # pragma: no cover - entrada manual
    ui.run(reload=False)

if __name__ in {"__main__", "__mp_main__"}:  # pragma: no cover
    main()
