Public Class Class1
    Public Class TestB
        Inherits Button
        Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
            MyBase.OnPaint(e)
            Dim f As New Drawing.Font("Arial", 9, FontStyle.Bold)
            Dim g As New Drawing.Font("Arial", 6, FontStyle.Bold)
            'Dim h As New Drawing.Font("Arial", 3, FontStyle.Bold)

            Dim elt() As String = {"H", "He", "Li", "Be", "B", "C", "N", "O", "F", "Ne", "Na", "Mg", "Al", "Si", "P", "S", "Cl", "Ar", "K", "Ca", "Sc", "Ti", "V", "Cr", "Mn", "Fe", "Co", "Ni", "Cu", "Zn", "Ga", "Ge", "As", "Se", "Br", "Kr", "Rb", "Sr", "Y", "Zr", "Nb", "Mo", "Tc", "Ru", "Rh", "Pd", "Ag", "Cd", "In", "Sn", "Sb", "Te", "I", "Xe", "Cs", "Ba", "La", "Ce", "Pr", "Nd", "Pm", "Sm", "Eu", "Gd", "Tb", "Dy", "Ho", "Er", "Tm", "Yb", "Lu", "Hf", "Ta", "W", "Re", "Os", "Ir", "Pt", "Au", "Hg", "Tl", "Pb", "Bi", "Po", "At", "Rn", "Fr", "Ra", "Ac", "Th", "Pa", "U", "Np", "Pu", "Am", "Cm", "Bk", "Cf", "Es"}
            'Dim element() As String = {"Hydrogen", "Helium", "Lithium", "Beryllium", "Boron", "Carbon", "Nitrogen", "Oxygen", "Fluorine", "Neon", "Sodium", "Magnesium", "Aluminum", "Silicon", "Phosphorus", "Sulfur", "Chlorine", "Argon", "Potassium", "Calcium", "Scandium", "Titanium", "Vanadium", "Chromium", "Manganese", "Iron", "Cobalt", "Nickel", "Copper", "Zinc", "Gallium", "Germanium", "Arsenic", "Selenium", "Bromine", "Krypton", "Rubidium", "Strontium", "Yttrium", "Zirconium", "Niobium", "Molybdenum", "Technetium", "Ruthenium", "Rhodium", "Palladium", "Silver", "Cadmium", "Indium", "Tin", "Antimony", "Tellurium", "Iodine", "Xenon", "Cesium", "Barium", "Lanthanum", "Cerium", "Praseodymium", "Neodymium", "Promethium", "Samarium", "Europium", "Gadolinium", "Terbium", "Dysprosium", "Holmium", "Erbium", "Thulium", "Ytterbium", "Lutetium", "Hafnium", "Tantalum", "Tungsten", "Rhenium", "Osmium", "Iridium", "Platinum", "Gold", "Mercury", "Thallium", "Lead", "Bismuth", "Polonium", "Astatine", "Radon", "Francium", "Radium", "Actinium", "Thorium", "Protactinium", "Uranium", "Neptunium", "Plutonium", "Americium", "Curium", "Berkelium", "Californium", "Einsteinium"}
            Dim ind As Integer = 0

            Dim tmp() As String = Me.Name.Split("t")
            If tmp.Length < 1 Then
                ReDim tmp(1)
                tmp(1) = 1
            End If
            If IsNumeric(tmp(1)) = False Then
                tmp(1) = 1
            End If

            Dim drawRect As New RectangleF(0, 2, Me.Width, Me.Height)
            'Dim drawRectBas As New RectangleF(-3, 0, 33, 24)

            Dim drawFormatCenter As New StringFormat
            drawFormatCenter.Alignment = StringAlignment.Center
            drawFormatCenter.LineAlignment = StringAlignment.Center

            'Dim drawFormat As New StringFormat
            'drawFormat.Alignment = StringAlignment.Center
            'drawFormat.LineAlignment = StringAlignment.Far

            If tmp(0) = "xray" Then
                e.Graphics.DrawString(tmp(1), f, Brushes.Black, drawRect, drawFormatCenter)
            Else
                e.Graphics.DrawString(elt(tmp(1) - 1), f, Brushes.Black, drawRect, drawFormatCenter)
                e.Graphics.DrawString(tmp(1), g, Brushes.Black, 3, 3)
            End If


            'e.Graphics.DrawString(element(tmp(1) - 1), h, Brushes.Black, drawRectBas, drawFormat)

        End Sub
    End Class
End Class
