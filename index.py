from flask import Flask, request, jsonify, render_template_string
from flask_cors import CORS
import win32print
import json
from datetime import datetime
import threading
import os
import sys
import tempfile
import atexit

app = Flask(__name__)
CORS(app)

def resource_path(relative_path):
    """Obtiene la ruta real del recurso tanto en desarrollo como compilado"""
    try:
        base_path = sys._MEIPASS  # PyInstaller temp folder
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Obtener el directorio temporal para los templates
def get_template_dir():
    if getattr(sys, 'frozen', False):
        # Estamos ejecutando como un ejecutable
        base_dir = sys._MEIPASS
    else:
        # Estamos ejecutando en desarrollo
        base_dir = os.path.dirname(os.path.abspath(__file__))

    template_dir = os.path.join(base_dir, 'templates')
    if not os.path.exists(template_dir):
        try:
            os.makedirs(template_dir)
        except PermissionError:
            # Si no tenemos permisos, usar directorio temporal
            template_dir = os.path.join(tempfile.gettempdir(), 'ImpresorTicketsMSM', 'templates')
            if not os.path.exists(template_dir):
                os.makedirs(template_dir, exist_ok=True)

    return template_dir


# Configurar la ruta de templates
template_dir = get_template_dir()
app.template_folder = template_dir


class UniversalThermalPrinter:
    def __init__(self, printer_name=None):
        self.printer_name = printer_name
        self.hprinter = None
        self.printer_type = "generic"

    def connect(self):
        try:
            if not self.printer_name:
                return False
            self.hprinter = win32print.OpenPrinter(self.printer_name)
            self._detect_printer_type()
            return True
        except Exception as e:
            print(f"Error conectando a impresora: {e}")
            return False

    def _detect_printer_type(self):
        printer_name_lower = self.printer_name.lower()
        escpos_brands = ['epson', 'star', 'bematech', 'daruma', 'zjiang', 'citizen']
        text_printers = ['generic', 'text', 'pdf', 'xps', 'microsoft print to pdf']

        if any(brand in printer_name_lower for brand in escpos_brands):
            self.printer_type = "escpos"
        elif any(text_printer in printer_name_lower for text_printer in text_printers):
            self.printer_type = "text_only"
        else:
            self.printer_type = "generic"

    def send_raw_data(self, data):
        try:
            if not self.hprinter:
                if not self.connect():
                    return False

            job_info = ("Python Print Job", None, "RAW")
            job_id = win32print.StartDocPrinter(self.hprinter, 1, job_info)
            win32print.StartPagePrinter(self.hprinter)

            if isinstance(data, str):
                data = data.encode('cp850', errors='replace')
            win32print.WritePrinter(self.hprinter, data)

            win32print.EndPagePrinter(self.hprinter)
            win32print.EndDocPrinter(self.hprinter)
            return True

        except Exception as e:
            print(f"Error enviando datos: {e}")
            return False

    def _safe_encode(self, text):
        try:
            return text.encode('cp850')
        except UnicodeEncodeError:
            try:
                return text.encode('cp437', errors='replace')
            except:
                safe_text = (text
                             .replace('Á', 'A').replace('á', 'a')
                             .replace('É', 'E').replace('é', 'e')
                             .replace('Í', 'I').replace('í', 'i')
                             .replace('Ó', 'O').replace('ó', 'o')
                             .replace('Ú', 'U').replace('ú', 'u')
                             .replace('Ñ', 'N').replace('ñ', 'n')
                             .replace('Ü', 'U').replace('ü', 'u')
                             .replace('¡', '').replace('¿', '')
                             )
                return safe_text.encode('cp437')

    def _get_printer_commands(self):
        if self.printer_type == "escpos":
            return {
                'init': b'\x1B\x40',  # Initialize printer (ESC @) - SIN saltos de línea
                'center': b'\x1B\x61\x01',
                'left': b'\x1B\x61\x00',
                'bold_on': b'\x1B\x45\x01',
                'bold_off': b'\x1B\x45\x00',
                'double_height': b'\x1B\x21\x10',
                'double_width_height': b'\x1B\x21\x30',
                'normal_text': b'\x1B\x21\x00',
                # Comandos de corte
                'cut_full': b'\x1D\x56\x00',
                'cut_partial': b'\x1D\x56\x01',
                'cut_feed_cut': b'\x1D\x56\x41\x00',
                'cut_epson': b'\x1D\x56\x01',
                'cut_star': b'\x1B\x64\x02',
                'cut_bematech': b'\x1B\x69',
                'cut_daruma': b'\x1B\x69',
                'cut_universal': b'\x1D\x56\x01',
                'feed_lines': b''  # SIN saltos de línea iniciales
            }
        elif self.printer_type == "text_only":
            return {
                'init': b'',  # Sin inicialización
                'center': b'',
                'left': b'',
                'bold_on': b'',
                'bold_off': b'',
                'double_height': b'',
                'double_width_height': b'',
                'normal_text': b'',
                'cut_full': b'\x0C',
                'cut_partial': b'\n',
                'cut_universal': b'\n',
                'feed_lines': b''  # SIN saltos de línea iniciales
            }
        else:
            return {
                'init': b'',  # Sin inicialización
                'center': b'',
                'left': b'',
                'bold_on': b'',
                'bold_off': b'',
                'double_height': b'',
                'double_width_height': b'',
                'normal_text': b'',
                'cut_full': b'\x1D\x56\x01',
                'cut_partial': b'\x1D\x56\x01',
                'cut_universal': b'\x1D\x56\x01',
                'feed_lines': b''  # SIN saltos de línea iniciales
            }

    # === FUNCIONES DE CENTRADO PERSONALIZADAS ===
    def _center_title(self, text):
        """Título principal (Mina San Miguel) - Grande, centrado y en negritas"""
        large_cmd = "\x1B\x21\x30"  # Doble altura y ancho
        normal_cmd = "\x1B\x21\x00"
        bold_on = "\x1B\x45\x01"
        bold_off = "\x1B\x45\x00"
        spaces = (24 - len(text)) // 2  # centrado a 28 cols (ajustado a tu título)
        return f"{large_cmd}{bold_on}{' ' * max(0, spaces)}{text}{bold_off}{normal_cmd}"

    def _center_subtitle(self, text):
        """Subtítulos (Sistema de Carga, Vale de Carga en Planta)"""
        bold_on = "\x1B\x45\x01"
        bold_off = "\x1B\x45\x00"
        spaces = (46 - len(text)) // 2
        return f"{bold_on}{' ' * max(0, spaces)}{text}{bold_off}"

    def _center_thanks(self, text):
        """Mensaje de gracias"""
        bold_on = "\x1B\x45\x01"
        bold_off = "\x1B\x45\x00"
        spaces = (48 - len(text)) // 2
        return f"{bold_on}{' ' * max(0, spaces)}{text}{bold_off}"

    def _center_footer(self, text):
        """Pie de página (Mina San Miguel - Confianza y Calidad)"""
        bold_on = "\x1B\x45\x01"
        bold_off = "\x1B\x45\x00"
        spaces = (50 - len(text)) // 2
        return f"{bold_on}{' ' * max(0, spaces)}{text}{bold_off}"

    def _center_tel(self, text):
        """Teléfono al pie"""
        bold_on = "\x1B\x45\x01"
        bold_off = "\x1B\x45\x00"
        spaces = (48 - len(text)) // 2
        return f"{bold_on}{' ' * max(0, spaces)}{text}{bold_off}"

    # === NUEVA FUNCIÓN ALINEADA POR COLUMNAS ===
    def _format_label_value(self, label, value, col_start=20, total_width=48):

        bold_on = "\x1B\x45\x01"  # ESC E 1 (negritas ON)
        bold_off = "\x1B\x45\x00"  # ESC E 0 (negritas OFF)

        # Texto de la etiqueta
        label_text = f"{bold_on}{label}:{bold_off}"

        # Calculamos cuántos espacios faltan para que el valor arranque en col_start
        spaces_needed = max(1, col_start - len(label) - 1)  # -1 por los dos puntos (:)
        spaces = " " * spaces_needed

        # Construimos la línea con el valor alineado
        line = f"{label_text}{spaces}{value}"

        # Ajustamos a total_width para que no se descuadre
        return line[:total_width] + "\n"

    def _qr_code(self, data: str, size: int = 6):
        """
        Genera UN SOLO código QR real para el ticket (ESC/POS correcto)
        """
        commands = []

        if not data or data.strip() == "":
            return commands

        try:
            # ---- ESCALAR +30% ----
            scaled_size = round(size * 1.3)
            size = max(3, min(12, scaled_size))

            store_len = len(data.encode('utf-8')) + 3
            pL = store_len % 256
            pH = store_len // 256

            commands.append(b'\n')
            commands.append(b'\x1B\x61\x01')  # centrar

            # limpiar buffer QR previo
            commands.append(b'\x1D\x28\x6B\x03\x00\x31\x52\x30')

            # configurar QR
            commands.append(b'\x1D\x28\x6B\x04\x00\x31\x41\x32\x00')  # modelo 2
            commands.append(bytes([0x1D, 0x28, 0x6B, 0x03, 0x00, 0x31, 0x43, size]))  # tamaño
            commands.append(b'\x1D\x28\x6B\x03\x00\x31\x45\x31')  # corrección M

            # guardar datos
            commands.append(
                bytes([0x1D, 0x28, 0x6B, pL, pH, 0x31, 0x50, 0x30]) +
                data.encode('utf-8')
            )

            # imprimir una sola vez
            commands.append(b'\x1D\x28\x6B\x03\x00\x31\x51\x30')

            # texto inferior
            commands.append(b'\n')

            commands.append(b'\x1B\x61\x00')

            return commands

        except Exception as e:
            print("QR error:", e)
            return [
                b'\n',
                b'\x1B\x61\x01',
                self._safe_encode(f"GITTICKET: {data}\n"),
                b'\x1B\x61\x00'
            ]

    def _generate_text_qr(self, data: str, size: int = 10):
        """
        Genera una representación visual del QR usando caracteres de bloque
        """
        import hashlib

        # Usar el dato para generar un patrón único
        hash_obj = hashlib.md5(data.encode())
        hash_hex = hash_obj.hexdigest()

        # Convertir a binario para crear el patrón
        binary = ''
        for char in hash_hex:
            binary += format(int(char, 16), '04b')

        # Asegurar que tenemos suficientes bits
        while len(binary) < size * size:
            binary += binary

        # Caracteres para los módulos del QR
        # Usamos bloques llenos y vacíos
        full_block = '█'
        empty_block = '░'

        lines = []
        for i in range(size):
            line = ''
            for j in range(size):
                idx = (i * size + j) % len(binary)
                line += full_block if binary[idx] == '1' else empty_block
            lines.append(line)

        return lines

    def _generate_barcode_fallback(self, data: str):
        """
        Genera un código de barras en texto usando caracteres
        """
        commands = []
        try:
            commands.append(b'\x1B\x61\x01')  # Centrar

            # Crear barras usando caracteres
            bars = []
            for char in data:
                # Convertir cada carácter a un patrón de barras
                val = ord(char) % 10
                bars.append('|' + '█' * (val + 1) + ' ')

            bar_line = ''.join(bars)
            commands.append(self._safe_encode(bar_line[:48] + '\n'))

            # Números debajo de las barras
            num_line = ''
            for char in data:
                num_line += f'{char}  '
            commands.append(self._safe_encode(num_line[:48] + '\n'))

            commands.append(b'\x1B\x61\x00')

        except:
            pass

        return commands

    # === FUCION PRINCIPAL PARA CONVERTIR IMAGEN A ESCALA DE GRISES
    def _image_to_escpos(self, image_path, width_bytes=12):  # Aumentado a 12 bytes = 96 píxeles
        """
        Convierte una imagen a comandos ESC/POS para impresora térmica
        width_bytes: ancho en bytes (cada byte = 8 pixeles)
        """
        try:
            from PIL import Image
            import math

            # Abrir y convertir imagen
            img = Image.open(image_path)

            # Convertir a escala de grises
            img = img.convert('L')

            # Redimensionar manteniendo proporción
            target_width = width_bytes * 8  # Ancho en píxeles
            aspect_ratio = img.height / img.width
            target_height = int(target_width * aspect_ratio)

            # Limitar altura máxima pero permitir más grande
            target_height = min(target_height, 160)  # Aumentado a 160 píxeles

            print(f"Redimensionando logo a: {target_width}x{target_height} píxeles")
            img = img.resize((target_width, target_height), Image.Resampling.LANCZOS)

            # Aplicar umbral para blanco y negro (mejorar contraste)
            threshold = 120  # Reducido para hacer más sensible a tonos oscuros
            img_bw = img.point(lambda x: 0 if x < threshold else 255, '1')

            commands = []

            # Comando GS v 0 para imprimir imagen raster
            xl = target_width // 8  # Ancho en bytes
            xh = 0
            yl = target_height % 256
            yh = target_height // 256

            print(f"Parámetros imagen: xl={xl}, xh={xh}, yl={yl}, yh={yh}")

            # Usar el modo que permite centrado con ESC/POS (modo 0 o 48)
            commands.append(b'\x1d\x76\x30\x00')  # Modo raster normal
            commands.append(bytes([xl, xh, yl, yh]))

            # Convertir imagen a datos binarios
            pixels = img_bw.load()
            for y in range(target_height):
                for x in range(0, target_width, 8):
                    byte = 0
                    for b in range(8):
                        if x + b < target_width:
                            # Pixel negro = bit 1, blanco = bit 0
                            if pixels[x + b, y] == 0:  # Negro
                                byte |= (1 << (7 - b))
                    commands.append(bytes([byte]))

            return b''.join(commands)

        except ImportError:
            print("Error: PIL no está instalado. Instálalo con: pip install Pillow")
            return b''
        except Exception as e:
            print(f"Error procesando imagen: {e}")
            import traceback
            traceback.print_exc()
            return b''

    def _print_logo(self, logo_path, width_bytes=12):
        """
        Imprime un logo centrado usando comandos ESC/POS reales
        width_bytes: ancho en bytes (12 bytes = 96 píxeles)
        """
        commands = []

        # Obtener los comandos de la impresora
        cmd = self._get_printer_commands()

        # Guardar estado actual de alineación
        # Comando ESC a n para alineación (0=izquierda, 1=centro, 2=derecha)
        align_center = b'\x1B\x61\x01'
        align_left = b'\x1B\x61\x00'

        # Centrar usando comando ESC/POS real
        commands.append(align_center)

        # Agregar imagen
        img_data = self._image_to_escpos(logo_path, width_bytes=width_bytes)
        if img_data and len(img_data) > 0:
            commands.append(img_data)
            commands.append(b'\n')  # Salto de línea después del logo
            print(f"Logo agregado exitosamente, tamaño: {len(img_data)} bytes")
        else:
            print("Error: No se pudo cargar la imagen del logo")
            # Fallback a texto
            commands.append(self._safe_encode("      [LOGO]      \n"))

        # Restaurar alineación izquierda
        commands.append(align_left)

        return commands

    # === FUNCIÓN PRINCIPAL DE IMPRESIÓN ===
    def print_ticket(self, ticket_data):
        try:
            commands = []
            cmd = self._get_printer_commands()

            # Inicializar impresora
            commands.append(cmd['init'])

            # IMPRIMIR LOGO MÁS GRANDE Y CENTRADO
            logo_path = resource_path("logo/Logo2.png")
            print(f"Intentando cargar logo desde: {logo_path}")

            # Usar width_bytes=14 para un logo aún más grande (112 píxeles de ancho)
            logo_commands = self._print_logo(logo_path, width_bytes=14)
            commands.extend(logo_commands)

            # Espacio después del logo
            commands.append(b'\n')

            # Título principal - MINA SAN MIGUEL (centrado y GRANDE)
            centered_title = self._center_title("MINA SAN MIGUEL")
            commands.append(self._safe_encode(f"{centered_title}\n"))

            # Subtítulo - SISTEMA DE CARGA
            centered_subtitle = self._center_subtitle("SISTEMA DE CARGA")
            commands.append(self._safe_encode(f"{centered_subtitle}\n"))

            # Separador
            commands.append(self._safe_encode("=" * 48 + "\n"))

            # Tipo de ticket
            centered_ticket_type = self._center_subtitle(ticket_data['TituloSecundario'])
            commands.append(self._safe_encode(f"{centered_ticket_type}\n"))
            commands.append(self._safe_encode("=" * 48 + "\n"))

            # Información básica
            commands.append(self._safe_encode(self._format_label_value("FECHA", ticket_data['Fecha'])))
            commands.append(self._safe_encode(self._format_label_value("FOLIO", ticket_data['Folio'])))

            # INFORMACION DE CARGA
            commands.append(self._safe_encode("-" * 48 + "\n"))
            commands.append(self._safe_encode("INFORMACION DE CARGA\n"))
            commands.append(self._safe_encode(self._format_label_value("PLANTA", ticket_data['NombrePlanta'])))
            commands.append(self._safe_encode(self._format_label_value("MATERIAL", ticket_data['NombreMaterial'])))
            commands.append(self._safe_encode(self._format_label_value("CLIENTE", ticket_data['NombreCliente'])))

            # DATOS DE TRANSPORTE
            commands.append(self._safe_encode("-" * 48 + "\n"))
            commands.append(self._safe_encode("DATOS DE TRANSPORTE\n"))
            commands.append(self._safe_encode(self._format_label_value("TRANSPORTE", ticket_data['Transporte'])))
            commands.append(self._safe_encode(self._format_label_value("PLACA", ticket_data['Placa'])))
            commands.append(self._safe_encode(self._format_label_value("RFID", ticket_data['RFID'])))

            # INFORMACION FINANCIERA
            commands.append(self._safe_encode("-" * 48 + "\n"))
            commands.append(self._safe_encode("INFORMACION FINANCIERA\n"))
            commands.append(self._safe_encode(self._format_label_value("CANTIDAD", ticket_data['Cantidad'])))
            commands.append(
                self._safe_encode(self._format_label_value("PRECIO UNIT", f"${ticket_data['PrecioUnidad']}")))
            commands.append(self._safe_encode(self._format_label_value("TOTAL", f"${ticket_data['TotalPago']}")))
            commands.append(self._safe_encode(self._format_label_value("FORMA PAGO", ticket_data['FormaPago'])))
            commands.append(self._safe_encode(self._format_label_value("VENDEDOR", ticket_data['Vendedor'])))

            # LLAMADA ÚNICA al QR
            qr_commands = self._qr_code(ticket_data['GitTicket'])
            commands.extend(qr_commands)  # <-- IMPORTANTE: usar extend, no append

            # Separador final (sin otro QR)
            commands.append(self._safe_encode("=" * 48 + "\n"))

            # Mensajes finales
            commands.append(self._safe_encode(f"{self._center_thanks('GRACIAS POR SU COMPRA!')}\n"))
            commands.append(self._safe_encode(f"{self._center_footer('Mina San Miguel - Confianza y Calidad')}\n"))
            commands.append(self._safe_encode(f"{self._center_tel('Tel: +52 722-492-5778')}\n"))

            # Alimentar papel
            commands.append(b'\n\n\n\n')

            # Corte
            self._perform_cut(commands, cmd)

            raw_data = b''.join(commands)
            return self.send_raw_data(raw_data)

        except Exception as e:
            print(f"Error en formato de ticket: {e}")
            import traceback
            traceback.print_exc()
            return False

    def _perform_cut(self, commands, cmd):
        """Realiza el corte del ticket sin espacios adicionales"""
        try:
            # SOLO el comando de corte, sin saltos de línea adicionales
            if self.printer_type == "escpos":
                commands.append(cmd['cut_partial'])  # Solo corte parcial
            elif self.printer_type == "text_only":
                commands.append(b'\x0C')  # Solo form feed
            else:
                commands.append(cmd['cut_universal'])  # Solo corte universal

        except Exception as e:
            print(f"Error en comando de corte: {e}")
            # Fallback mínimo
            commands.append(b'\n')  # Solo un salto de línea como fallback

    def print_test_receipt(self):
        try:
            commands = []
            cmd = self._get_printer_commands()

            commands.append(cmd['init'])

            if cmd['center']:
                commands.append(cmd['center'])

            commands.append(self._safe_encode("INDUSTRIAS MACON\n"))
            commands.append(self._safe_encode("Sistema de Carga Automatizado\n"))
            commands.append(self._safe_encode("=" * 32 + "\n"))

            if cmd['left']:
                commands.append(cmd['left'])

            commands.append(self._safe_encode("RECIBO DE PRUEBA\n"))
            commands.append(self._safe_encode(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"))
            commands.append(self._safe_encode("-" * 32 + "\n"))

            items = [
                ("Carga Material A", "150.00 TON"),
                ("Manejo Especial", "25.50"),
                ("Servicio Premium", "75.00")
            ]

            for item, price in items:
                commands.append(self._safe_encode(f"{item:<20} ${price:>10}\n"))

            commands.append(self._safe_encode("-" * 32 + "\n"))
            commands.append(self._safe_encode("TOTAL: $250.50\n"))
            commands.append(self._safe_encode("=" * 32 + "\n"))

            if cmd['center']:
                commands.append(cmd['center'])
            commands.append(self._safe_encode("Sistema funcionando correctamente\n"))

            # Alimentar papel antes del corte
            commands.append(b'\n\n\n')

            # Realizar el corte
            self._perform_cut(commands, cmd)

            raw_data = b''.join(commands)
            return self.send_raw_data(raw_data)

        except Exception as e:
            print(f"Error en test: {e}")
            import traceback
            traceback.print_exc()
            return False

    def test_cut(self):
        """Prueba específica del mecanismo de corte"""
        try:
            if not self.hprinter:
                if not self.connect():
                    return False

            commands = []
            cmd = self._get_printer_commands()

            commands.append(cmd['init'])
            commands.append(self._safe_encode("PRUEBA DE CORTE\n"))
            commands.append(self._safe_encode("Este ticket debe cortarse automaticamente\n"))
            commands.append(b'\n\n\n\n\n')

            # Probar el corte
            self._perform_cut(commands, cmd)

            raw_data = b''.join(commands)
            return self.send_raw_data(raw_data)

        except Exception as e:
            print(f"Error en prueba de corte: {e}")
            return False

    def test_impresion_nativa(self):
        """
        Realiza una prueba nativa completa de la impresora seleccionada
        mostrando toda la información disponible y probando todas sus capacidades
        """
        try:
            if not self.hprinter:
                if not self.connect():
                    return False

            commands = []
            cmd = self._get_printer_commands()

            # Inicializar impresora
            commands.append(cmd['init'])

            # Encabezado de la prueba
            commands.append(self._safe_encode("=" * 48 + "\n"))
            commands.append(cmd['center'])
            commands.append(cmd['double_width_height'])
            commands.append(self._safe_encode("INFORME COMPLETO DE IMPRESORA\n"))
            commands.append(cmd['normal_text'])
            commands.append(cmd['left'])
            commands.append(self._safe_encode("=" * 48 + "\n\n"))

            # Información básica de la impresora
            commands.append(cmd['bold_on'])
            commands.append(self._safe_encode("INFORMACIÓN BÁSICA:\n"))
            commands.append(cmd['bold_off'])
            commands.append(self._safe_encode(f"• Nombre: {self.printer_name}\n"))
            commands.append(self._safe_encode(f"• Tipo detectado: {self.printer_type}\n"))
            commands.append(self._safe_encode(f"• Fecha/Hora: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"))

            # Obtener información detallada de la impresora usando win32print
            try:
                printer_info = win32print.GetPrinter(self.hprinter, 2)
                commands.append(self._safe_encode(f"• Estado: {printer_info['Status']}\n"))
                commands.append(self._safe_encode(f"• Servidor: {printer_info['pServerName']}\n"))
                commands.append(self._safe_encode(f"• Puerto: {printer_info['pPortName']}\n"))
                commands.append(self._safe_encode(f"• Controlador: {printer_info['pDriverName']}\n"))

                # Intentar obtener más detalles
                attributes = printer_info['Attributes']
                attr_details = []
                if attributes & win32print.PRINTER_ATTRIBUTE_LOCAL:
                    attr_details.append("Local")
                if attributes & win32print.PRINTER_ATTRIBUTE_NETWORK:
                    attr_details.append("Red")
                if attributes & win32print.PRINTER_ATTRIBUTE_SHARED:
                    attr_details.append("Compartida")
                if attributes & win32print.PRINTER_ATTRIBUTE_DIRECT:
                    attr_details.append("Directa")

                commands.append(
                    self._safe_encode(f"• Atributos: {', '.join(attr_details) if attr_details else 'Ninguno'}\n"))

            except Exception as info_error:
                commands.append(self._safe_encode(f"• Error obteniendo detalles: {str(info_error)}\n"))

            commands.append(self._safe_encode("\n"))

            # CAPACIDADES DE LA IMPRESORA
            commands.append(cmd['bold_on'])
            commands.append(self._safe_encode("CAPACIDADES DETECTADAS:\n"))
            commands.append(cmd['bold_off'])

            if self.printer_type == "escpos":
                commands.append(self._safe_encode("• Soporte ESC/POS: COMPLETO\n"))
                commands.append(self._safe_encode("• Formatos de texto: Negrita, Centrado, Tamaños\n"))
                commands.append(self._safe_encode("• Códigos de barras: Soportados\n"))
                commands.append(self._safe_encode("• Corte de papel: Automático\n"))
            elif self.printer_type == "text_only":
                commands.append(self._safe_encode("• Tipo: Solo texto\n"))
                commands.append(self._safe_encode("• Formatos: Básicos (sin especiales)\n"))
                commands.append(self._safe_encode("• Corte de papel: Por form feed (0x0C)\n"))
            else:
                commands.append(self._safe_encode("• Tipo: Genérica/Desconocida\n"))
                commands.append(self._safe_encode("• Formatos: Limitados o desconocidos\n"))
                commands.append(self._safe_encode("• Corte de papel: Intento universal\n"))

            commands.append(self._safe_encode("\n"))

            # PRUEBA DE CARACTERES Y CODIFICACIÓN
            commands.append(cmd['bold_on'])
            commands.append(self._safe_encode("PRUEBA DE CARACTERES:\n"))
            commands.append(cmd['bold_off'])

            # Caracteres ASCII estándar
            commands.append(self._safe_encode("ASCII: !\"#$%&'()*+,-./0123456789:;<=>?@\n"))
            commands.append(self._safe_encode("ASCII: ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`\n"))
            commands.append(self._safe_encode("ASCII: abcdefghijklmnopqrstuvwxyz{|}~\n"))

            # Caracteres extendidos (dependiendo de la codificación)
            commands.append(self._safe_encode("Extendidos: áéíóúñÑüÜ¿¡°ªº€§¶©®™\n"))

            # Caracteres especiales para impresoras térmicas
            commands.append(self._safe_encode("Especiales: █▓▒░┌┐└┘├┤┬┴┼═║╔╗╚╝╠╣╦╩╬\n"))

            commands.append(self._safe_encode("\n"))

            # PRUEBA DE FORMATOS (solo para ESC/POS)
            if self.printer_type == "escpos":
                commands.append(cmd['bold_on'])
                commands.append(self._safe_encode("PRUEBA DE FORMATOS ESC/POS:\n"))
                commands.append(cmd['bold_off'])

                # Texto normal
                commands.append(self._safe_encode("Texto normal - "))

                # Texto en negrita
                commands.append(cmd['bold_on'])
                commands.append(self._safe_encode("Negrita"))
                commands.append(cmd['bold_off'])
                commands.append(self._safe_encode(" - "))

                # Doble altura
                commands.append(cmd['double_height'])
                commands.append(self._safe_encode("Doble altura"))
                commands.append(cmd['normal_text'])
                commands.append(self._safe_encode(" - "))

                # Doble altura y ancho
                commands.append(cmd['double_width_height'])
                commands.append(self._safe_encode("Doble tamaño"))
                commands.append(cmd['normal_text'])
                commands.append(self._safe_encode("\n"))

                # Texto centrado
                commands.append(cmd['center'])
                commands.append(self._safe_encode("Texto centrado\n"))
                commands.append(cmd['left'])

                # Subrayado (si está disponible)
                try:
                    commands.append(b'\x1B\x2D\x01')  # Subrayado simple ON
                    commands.append(self._safe_encode("Texto subrayado"))
                    commands.append(b'\x1B\x2D\x00')  # Subrayado OFF
                    commands.append(self._safe_encode("\n"))
                except:
                    commands.append(self._safe_encode("(Subrayado no disponible)\n"))

                commands.append(self._safe_encode("\n"))

            # PRUEBA DE CÓDIGOS DE BARRAS (solo para ESC/POS)
            if self.printer_type == "escpos":
                commands.append(cmd['bold_on'])
                commands.append(self._safe_encode("PRUEBA DE CÓDIGOS DE BARRAS:\n"))
                commands.append(cmd['bold_off'])

                try:
                    # CODE39
                    commands.append(self._safe_encode("CODE39: "))
                    barcode_data = "TEST123"
                    barcode_cmd = b'\x1D\x6B\x04' + barcode_data.encode('ascii') + b'\x00'
                    commands.append(barcode_cmd)
                    commands.append(self._safe_encode("\n"))

                    # UPC-A (si el dato tiene 11-12 dígitos)
                    commands.append(self._safe_encode("UPC-A: "))
                    upc_data = "12345678901"  # 11 dígitos para UPC-A
                    upc_cmd = b'\x1D\x6B\x00' + upc_data.encode('ascii') + b'\x00'
                    commands.append(upc_cmd)
                    commands.append(self._safe_encode("\n"))

                except Exception as barcode_error:
                    commands.append(self._safe_encode(f"Error en códigos de barras: {barcode_error}\n"))

                commands.append(self._safe_encode("\n"))

            # INFORMACIÓN DEL SISTEMA
            commands.append(cmd['bold_on'])
            commands.append(self._safe_encode("INFORMACIÓN DEL SISTEMA:\n"))
            commands.append(cmd['bold_off'])

            try:
                import platform
                import socket

                commands.append(self._safe_encode(f"• Sistema: {platform.system()} {platform.release()}\n"))
                commands.append(self._safe_encode(f"• Python: {platform.python_version()}\n"))
                commands.append(self._safe_encode(f"• Hostname: {socket.gethostname()}\n"))
                commands.append(self._safe_encode(f"• Usuario: {os.getlogin()}\n"))

            except Exception as sys_error:
                commands.append(self._safe_encode(f"• Error info sistema: {sys_error}\n"))

            commands.append(self._safe_encode("\n"))

            # RESUMEN Y RECOMENDACIONES
            commands.append(cmd['bold_on'])
            commands.append(self._safe_encode("RESUMEN Y ESTADO:\n"))
            commands.append(cmd['bold_off'])

            commands.append(self._safe_encode("✓ Conexión establecida correctamente\n"))
            commands.append(self._safe_encode("✓ Comandos básicos funcionando\n"))

            if self.printer_type == "escpos":
                commands.append(self._safe_encode("✓ Impresora ESC/POS detectada\n"))
                commands.append(self._safe_encode("✓ Formatos avanzados disponibles\n"))
            elif self.printer_type == "text_only":
                commands.append(self._safe_encode("⚠ Impresora de solo texto\n"))
                commands.append(self._safe_encode("⚠ Formatos limitados\n"))
            else:
                commands.append(self._safe_encode("? Tipo de impresora desconocido\n"))
                commands.append(self._safe_encode("? Pruebe configuraciones manuales\n"))

            commands.append(self._safe_encode("\n"))

            # FIRMA Y FINAL
            commands.append(cmd['center'])
            commands.append(self._safe_encode("--- FIN DEL REPORTE ---\n"))
            commands.append(cmd['left'])

            # Alimentar papel suficiente antes del corte
            commands.append(b'\n\n\n\n\n')

            # Realizar corte según el tipo de impresora
            self._perform_cut(commands, cmd)

            # Enviar todos los comandos
            raw_data = b''.join(commands)
            return self.send_raw_data(raw_data)

        except Exception as e:
            print(f"Error en prueba nativa completa: {e}")
            import traceback
            traceback.print_exc()
            return False

    def close(self):
        if self.hprinter:
            win32print.ClosePrinter(self.hprinter)
            self.hprinter = None


printer = UniversalThermalPrinter()


def get_available_printers():
    printers = []
    try:
        printer_info = win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        )
        for printer in printer_info:
            printers.append({
                'name': printer[2],
                'description': printer[1] or printer[2],
                'status': printer[3]
            })
    except Exception as e:
        print(f"Error obteniendo impresoras: {e}")
    return printers


# HTML template como string para evitar problemas de archivo
TEMPLATE_HTML = '''
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistema de Impresión - Mina San Miguel</title>
    <style>
        /* Variables CSS */
        :root {
            --primary-color: #2c3e50;
            --secondary-color: #3498db;
            --accent-color: #e74c3c;
            --success-color: #27ae60;
            --warning-color: #f39c12;
            --light-color: #ecf0f1;
            --dark-color: #2c3e50;
            --text-color: #333;
            --border-radius: 8px;
            --box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            --transition: all 0.3s ease;
        }

        /* Reset y base */
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
            color: var(--text-color);
            line-height: 1.6;
        }

        .container {
            max-width: 1000px;
            margin: 0 auto;
            background: white;
            border-radius: var(--border-radius);
            box-shadow: var(--box-shadow);
            overflow: hidden;
        }

        /* Header */
        .header {
            background: var(--primary-color);
            color: white;
            padding: 25px;
            text-align: center;
            position: relative;
        }

        .header h1 {
            font-size: 2.2em;
            margin-bottom: 5px;
            font-weight: 300;
        }

        .header h1 strong {
            font-weight: 600;
        }

        .header p {
            opacity: 0.9;
            font-size: 1.1em;
        }

        .header::after {
            content: '';
            position: absolute;
            bottom: 0;
            left: 0;
            right: 0;
            height: 4px;
            background: linear-gradient(90deg, var(--secondary-color), var(--accent-color));
        }

        /* Contenido principal */
        .content {
            padding: 30px;
        }

        /* Tarjeta de estado */
        .status-card {
            background: var(--light-color);
            padding: 20px;
            border-radius: var(--border-radius);
            margin-bottom: 25px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 15px;
        }

        .printer-info {
            flex: 1;
            min-width: 250px;
        }

        .current-printer {
            font-size: 1.2em;
            font-weight: 600;
            color: var(--primary-color);
            margin-bottom: 5px;
        }

        .printer-type {
            font-size: 0.9em;
            color: #666;
            padding: 4px 8px;
            background: #e0e0e0;
            border-radius: 12px;
            display: inline-block;
        }

        .status-indicator {
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .status-dot {
            width: 12px;
            height: 12px;
            border-radius: 50%;
            background: var(--accent-color);
        }

        .status-dot.connected {
            background: var(--success-color);
        }

        /* Secciones */
        .section {
            margin-bottom: 30px;
        }

        .section-title {
            font-size: 1.4em;
            color: var(--primary-color);
            margin-bottom: 15px;
            padding-bottom: 10px;
            border-bottom: 2px solid var(--light-color);
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .section-title::before {
            content: '📋';
            font-size: 1.2em;
        }

        .section-title.actions::before {
            content: '⚙️';
        }

        /* Lista de impresoras */
        .printer-list {
            display: grid;
            gap: 12px;
            margin-bottom: 20px;
        }

        .printer-item {
            padding: 15px;
            border: 2px solid #e0e0e0;
            border-radius: var(--border-radius);
            cursor: pointer;
            transition: var(--transition);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .printer-item:hover {
            border-color: var(--secondary-color);
            transform: translateY(-2px);
            box-shadow: var(--box-shadow);
        }

        .printer-item.selected {
            border-color: var(--secondary-color);
            background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%);
        }

        .printer-details {
            flex: 1;
        }

        .printer-name {
            font-weight: 600;
            color: var(--primary-color);
            margin-bottom: 5px;
        }

        .printer-description {
            font-size: 0.9em;
            color: #666;
        }

        .printer-status {
            font-size: 0.8em;
            padding: 4px 8px;
            border-radius: 12px;
            background: #e0e0e0;
        }

        .printer-status.available {
            background: #d4edda;
            color: #155724;
        }

        /* Botones */
        .btn {
            padding: 12px 24px;
            border: none;
            border-radius: var(--border-radius);
            cursor: pointer;
            font-weight: 600;
            transition: var(--transition);
            display: inline-flex;
            align-items: center;
            gap: 8px;
            text-decoration: none;
            font-family: inherit;
        }

        .btn-primary {
            background: var(--secondary-color);
            color: white;
        }

        .btn-primary:hover {
            background: #2980b9;
            transform: translateY(-2px);
        }

        .btn-primary:disabled {
            background: #bdc3c7;
            cursor: not-allowed;
            transform: none;
        }

        .btn-success {
            background: var(--success-color);
            color: white;
        }

        .btn-success:hover {
            background: #229954;
            transform: translateY(-2px);
        }

        .btn-warning {
            background: var(--warning-color);
            color: white;
        }

        .btn-warning:hover {
            background: #e67e22;
            transform: translateY(-2px);
        }

        .btn-danger {
            background: var(--accent-color);
            color: white;
        }

        .btn-danger:hover {
            background: #c0392b;
            transform: translateY(-2px);
        }

        /* Grid de acciones */
        .actions-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin-top: 20px;
        }

        /* Mensajes de estado */
        .status-message {
            padding: 15px;
            border-radius: var(--border-radius);
            margin: 15px 0;
            display: none;
            animation: slideIn 0.3s ease;
        }

        .status-success {
            background: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }

        .status-error {
            background: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }

        @keyframes slideIn {
            from {
                opacity: 0;
                transform: translateY(-10px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }

        /* Footer */
        .footer {
            text-align: center;
            padding: 20px;
            background: var(--light-color);
            color: #666;
            font-size: 0.9em;
        }

        /* Loading animation */
        .loading {
            display: inline-block;
            width: 20px;
            height: 20px;
            border: 3px solid #f3f3f3;
            border-top: 3px solid var(--secondary-color);
            border-radius: 50%;
            animation: spin 1s linear infinite;
        }

        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }

        /* Responsive */
        @media (max-width: 768px) {
            .content {
                padding: 20px;
            }

            .status-card {
                flex-direction: column;
                align-items: flex-start;
            }

            .actions-grid {
                grid-template-columns: 1fr;
            }

            .printer-item {
                flex-direction: column;
                align-items: flex-start;
                gap: 10px;
            }

            .header h1 {
                font-size: 1.8em;
            }

            .section-title {
                font-size: 1.2em;
            }
        }

        @media (max-width: 480px) {
            body {
                padding: 10px;
            }

            .header {
                padding: 20px 15px;
            }

            .content {
                padding: 15px;
            }

            .btn {
                padding: 10px 16px;
                font-size: 0.9em;
            }
        }

        /* Estados visuales */
        .printer-item[data-status="0"] .printer-status {
            background: #d4edda;
            color: #155724;
        }

        .printer-item[data-status="1"] .printer-status {
            background: #f8d7da;
            color: #721c24;
        }

        /* Mejoras de accesibilidad */
        .btn:focus,
        .printer-item:focus {
            outline: 2px solid var(--secondary-color);
            outline-offset: 2px;
        }

        /* Efectos de profundidad */
        .container {
            position: relative;
            z-index: 1;
        }

        .container::before {
            content: '';
            position: absolute;
            top: -10px;
            left: -10px;
            right: -10px;
            bottom: -10px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            z-index: -1;
            filter: blur(20px);
            opacity: 0.3;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1><strong>Mina San Miguel</strong> - Sistema de Impresión</h1>
            <p>Sistema de Carga y Control de Tickets</p>
        </div>

        <div class="content">
            <div class="status-card">
                <div class="printer-info">
                    <div class="current-printer">
                        📋 Impresora actual: {{ current_printer or 'No seleccionada' }}
                    </div>
                    <div class="printer-type">Tipo: Esperando conexión...</div>
                </div>
                <div class="status-indicator">
                    <div class="status-dot"></div>
                    <span>Desconectado</span>
                </div>
            </div>

            <div class="section">
                <h2 class="section-title">Impresoras Disponibles</h2>
                <div class="printer-list" id="printerList">
                    {% for printer in printers %}
                    <div class="printer-item" data-name="{{ printer.name }}" data-status="{{ printer.status }}">
                        <div class="printer-details">
                            <div class="printer-name">{{ printer.name }}</div>
                            <div class="printer-description">{{ printer.description }}</div>
                        </div>
                        <div class="printer-status">
                            {% if printer.status == 0 %}Disponible{% else %}Ocupada{% endif %}
                        </div>
                    </div>
                    {% endfor %}
                </div>

                <button id="selectButton" class="btn btn-primary" disabled>
                    ✅ Seleccionar Impresora
                </button>

                <div class="status-message" id="statusMessage"></div>
            </div>

            <div class="section">
                <h2 class="section-title actions">Acciones del Sistema</h2>
                <div class="actions-grid">
                    <button class="btn btn-primary" onclick="testNativo()">
                        🖨️ Test de Impresion
                    </button>
                    <button class="btn btn-warning" onclick="testReceipt()">
                        📄 Test de Recibo
                    </button>
                    <button class="btn btn-danger" onclick="testCut()">
                        ✂️ Probar Corte
                    </button>
                    <button class="btn btn-primary" onclick="checkStatus()">
                        🔍 Verificar Estado
                    </button>
                    <button class="btn btn-primary" onclick="refreshPrinters()">
                        🔄 Actualizar Lista
                    </button>
                </div>
            </div>
        </div>

        <div class="footer">
            <p>© 2024 Mina San Miguel - Sistema de Carga Automatizado</p>
            <p>Versión 1.0 | Confianza y Calidad</p>
        </div>
    </div>

    <script>
        let selectedPrinter = null;
        let currentStatus = false;

        document.querySelectorAll('.printer-item').forEach(item => {
            item.addEventListener('click', () => {
                document.querySelectorAll('.printer-item').forEach(i => i.classList.remove('selected'));
                item.classList.add('selected');
                selectedPrinter = item.dataset.name;
                document.getElementById('selectButton').disabled = false;
            });
        });

        document.getElementById('selectButton').addEventListener('click', async () => {
            if (!selectedPrinter) return;

            const button = document.getElementById('selectButton');
            const originalText = button.innerHTML;
            button.innerHTML = '<span class="loading"></span> Procesando...';
            button.disabled = true;

            try {
                const response = await fetch('/select-printer', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ printer_name: selectedPrinter })
                });

                const data = await response.json();
                const statusDiv = document.getElementById('statusMessage');
                statusDiv.style.display = 'block';

                if (response.ok) {
                    statusDiv.className = 'status-message status-success';
                    statusDiv.innerHTML = '✅ ' + data.message;
                    document.querySelector('.current-printer').textContent = 
                        '📋 Impresora actual: ' + selectedPrinter;
                    document.querySelector('.printer-type').textContent = 'Tipo: ' + data.printer_type;
                    updateStatus(true);
                } else {
                    statusDiv.className = 'status-message status-error';
                    statusDiv.innerHTML = '❌ ' + data.error;
                }

                setTimeout(() => { 
                    statusDiv.style.display = 'none'; 
                    button.innerHTML = originalText;
                    button.disabled = false;
                }, 3000);
            } catch (error) {
                console.error('Error:', error);
                button.innerHTML = originalText;
                button.disabled = false;
            }
        });

        async function testPrint() {
            showLoading('Test de Impresión');
            try {
                const response = await fetch('/test-print');
                const data = await response.json();
                showAlert(response.ok, response.ok ? data.message : data.error);
            } catch (error) {
                showAlert(false, 'Error: ' + error.message);
            }
        }

        async function testNativo() {
            showLoading('Test Nativo');
            try {
                const response = await fetch('/test-nativo');
                const data = await response.json();
                showAlert(response.ok, response.ok ? data.message : data.error);
            } catch (error) {
                showAlert(false, 'Error: ' + error.message);
            }
        }

        async function testReceipt() {
            showLoading('Test de Recibo');
            try {
                const response = await fetch('/test-receipt');
                const data = await response.json();
                showAlert(response.ok, response.ok ? data.message : data.error);
            } catch (error) {
                showAlert(false, 'Error: ' + error.message);
            }
        }

        async function testCut() {
            showLoading('Prueba de Corte');
            try {
                const response = await fetch('/test-cut');
                const data = await response.json();
                showAlert(response.ok, response.ok ? data.message : data.error);
            } catch (error) {
                showAlert(false, 'Error: ' + error.message);
            }
        }

        async function checkStatus() {
            showLoading('Verificando Estado');
            try {
                const response = await fetch('/health');
                const data = await response.json();
                updateStatus(data.printer_connected);
                showAlert(true, 
                    `Estado: ${data.printer_connected ? 'Conectado' : 'Desconectado'}\n` +
                    `Tipo: ${data.printer_type}\n` +
                    `Seleccionada: ${data.printer_selected ? 'Sí' : 'No'}`
                );
            } catch (error) {
                showAlert(false, 'Error: ' + error.message);
            }
        }

        async function refreshPrinters() {
            showLoading('Actualizando');
            try {
                const response = await fetch('/get-printers');
                const data = await response.json();
                if (response.ok) {
                    location.reload();
                } else {
                    showAlert(false, 'Error al actualizar: ' + data.error);
                }
            } catch (error) {
                showAlert(false, 'Error: ' + error.message);
            }
        }

        function updateStatus(connected) {
            currentStatus = connected;
            const statusDot = document.querySelector('.status-dot');
            const statusText = document.querySelector('.status-indicator span');

            if (connected) {
                statusDot.className = 'status-dot connected';
                statusText.textContent = 'Conectado';
                statusText.style.color = 'var(--success-color)';
            } else {
                statusDot.className = 'status-dot';
                statusText.textContent = 'Desconectado';
                statusText.style.color = 'var(--accent-color)';
            }
        }

        function showLoading(action) {
            const statusDiv = document.getElementById('statusMessage');
            statusDiv.className = 'status-message status-success';
            statusDiv.innerHTML = '<span class="loading"></span> ' + action + '...';
            statusDiv.style.display = 'block';
        }

        function showAlert(success, message) {
            const statusDiv = document.getElementById('statusMessage');
            statusDiv.className = success ? 'status-message status-success' : 'status-message status-error';
            statusDiv.innerHTML = success ? 
                '✅ ' + message : 
                '❌ ' + message;
            statusDiv.style.display = 'block';

            setTimeout(() => { 
                statusDiv.style.display = 'none'; 
            }, 3000);
        }

        async function loadPrinters() {
            try {
                const response = await fetch('/get-printers');
                const data = await response.json();
                if (response.ok && data.current_printer) {
                    document.querySelector('.current-printer').textContent = 
                        '📋 Impresora actual: ' + data.current_printer;
                }

                const healthResponse = await fetch('/health');
                const healthData = await healthResponse.json();
                if (healthResponse.ok) {
                    updateStatus(healthData.printer_connected);
                    document.querySelector('.printer-type').textContent = 
                        'Tipo: ' + (healthData.printer_type || 'No detectado');
                }
            } catch (error) {
                console.error('Error cargando impresoras:', error);
            }
        }

        loadPrinters();
    </script>
</body>
</html>'''


@app.route('/')
def index():
    printers = get_available_printers()
    current_printer = printer.printer_name if printer.printer_name else "No seleccionada"
    return render_template_string(TEMPLATE_HTML,
                                  printers=printers,
                                  current_printer=current_printer)


@app.route('/select-printer', methods=['POST'])
def select_printer():
    try:
        data = request.get_json()
        printer_name = data.get('printer_name')

        if not printer_name:
            return jsonify({'error': 'No se especificó impresora'}), 400

        printer.printer_name = printer_name
        printer.close()

        if printer.connect():
            return jsonify({
                'message': f'Impresora seleccionada: {printer_name}',
                'printer_name': printer_name,
                'printer_type': printer.printer_type
            }), 200
        else:
            return jsonify({
                'error': f'No se pudo conectar a la impresora: {printer_name}'
            }), 500

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/get-printers', methods=['GET'])
def get_printers():
    try:
        printers = get_available_printers()
        current_printer = printer.printer_name if printer.printer_name else None

        return jsonify({
            'printers': printers,
            'current_printer': current_printer
        }), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/current-printer', methods=['GET'])
def current_printer():
    return jsonify({
        'current_printer': printer.printer_name if printer.printer_name else None
    }), 200


@app.route('/imprimir-ticket', methods=['POST'])
def imprimir_ticket():
    try:
        if not printer.printer_name:
            return jsonify({'error': 'No hay impresora seleccionada'}), 400

        data = request.get_json()
        if not data:
            return jsonify({'error': 'No se recibieron datos'}), 400

        print("Datos recibidos:", data)

        cleaned_data = {}
        for key, value in data.items():
            if isinstance(value, str):
                # Limpiar caracteres especiales (tu código original)
                cleaned_value = (value
                                 .replace('Á', 'A').replace('á', 'a')
                                 .replace('É', 'E').replace('é', 'e')
                                 .replace('Í', 'I').replace('í', 'i')
                                 .replace('Ó', 'O').replace('ó', 'o')
                                 .replace('Ú', 'U').replace('ú', 'u')
                                 .replace('Ñ', 'N').replace('ñ', 'n')
                                 .replace('Ü', 'U').replace('ü', 'u')
                                 .replace('¡', '').replace('¿', '')
                                 )

                # NUEVO: Limpiar prefijo 1Q específicamente para RFID
                if key == 'RFID' and cleaned_value.startswith('1Q'):
                    cleaned_value = cleaned_value[2:]
                    print(f"RFID limpiado: {cleaned_value}")

                cleaned_data[key] = cleaned_value
            else:
                cleaned_data[key] = value

        print("Datos limpios:", cleaned_data)
        success = printer.print_ticket(cleaned_data)

        if success:
            return jsonify({'message': 'Ticket impreso correctamente'}), 200
        else:
            return jsonify({'error': 'Error al imprimir'}), 500

    except Exception as e:
        print(f"Error en el servidor: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/health', methods=['GET'])
def health_check():
    status = printer.connect() if printer.printer_name else False
    return jsonify({
        'status': 'ok',
        'printer_connected': status,
        'printer_selected': printer.printer_name is not None,
        'printer_type': printer.printer_type
    })


@app.route('/test-print', methods=['GET'])
def test_print():
    try:
        if not printer.printer_name:
            return jsonify({'error': 'No hay impresora seleccionada'}), 400

        test_data = {
            'TituloSecundario': 'VALE DE CARGA - PLANTA',
            'Fecha': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'Folio': 'MAC-2024-00125',
            'NombrePlanta': 'Planta Central Norte',
            'NombreMaterial': 'Concreto Premezclado',
            'Cantidad': '150.75',
            'PrecioUnidad': '185.50',
            'TotalPago': '27,954.12',
            'FormaPago': 'Transferencia Electronica',
            'Transporte': 'Volvo FH16-750',
            'Placa': 'ABC-1234',
            'Vendedor': 'Juan Perez',
            'RFID': 'MAC-RFID-8876',
            'NombreCliente': 'Constructora Moderna SA de CV'
        }

        success = printer.print_ticket(test_data)
        if success:
            return jsonify({'message': 'Test impreso correctamente'}), 200
        else:
            return jsonify({'error': 'Error en test de impresión'}), 500

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/test-nativo', methods=['GET'])
def test_nativo():
    try:
        if not printer.printer_name:
            return jsonify({'error': 'No hay impresora seleccionada'}), 400

        success = printer.test_impresion_nativa()
        if success:
            return jsonify({'message': 'Prueba nativa ejecutada correctamente'}), 200
        else:
            return jsonify({'error': 'Error en prueba nativa de impresión'}), 500
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/test-receipt', methods=['GET'])
def test_receipt():
    try:
        if not printer.printer_name:
            return jsonify({'error': 'No hay impresora seleccionada'}), 400

        success = printer.print_test_receipt()
        if success:
            return jsonify({'message': 'Recibo de prueba impreso'}), 200
        else:
            return jsonify({'error': 'Error en recibo de prueba'}), 500
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/test-cut', methods=['GET'])
def test_cut():
    try:
        if not printer.printer_name:
            return jsonify({'error': 'No hay impresora seleccionada'}), 400

        success = printer.test_cut()
        if success:
            return jsonify({'message': 'Prueba de corte ejecutada'}), 200
        else:
            return jsonify({'error': 'Error en prueba de corte'}), 500
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


def run_server():
    print("Servidor iniciado en http://localhost:5000")
    print("Endpoints disponibles:")
    print("  GET  /              - Interfaz de selección de impresora")
    print("  GET  /health        - Estado de la impresora")
    print("  GET  /get-printers  - Lista de impresoras disponibles")
    print("  POST /select-printer- Seleccionar impresora")
    print("  GET  /test-print    - Test de ticket profesional")
    print("  GET  /test-receipt  - Test de recibo profesional")
    print("  GET  /test-cut      - Prueba de corte")
    print("  POST /imprimir-ticket - Imprimir ticket desde la app")

    # Ejecutar en modo sin depuración para producción
    app.run(host='0.0.0.0', port=5000, debug=False, threaded=True)


if __name__ == '__main__':
    run_server()