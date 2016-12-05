Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Imaging

Module ImageConverter

    '
    ' API
    '

    ' Convert input RGB image to Indexed image
    ' @param SrcFile - source file name or path.
    ' @param DstFile - destination file name or path. Destination file will be overwritten.
    ' @param NumberOfColors - number of colors in destination image

    Public Sub RGB_2_Index4(aSrcFileName As String, aDstFileName As String)

        Dim srcImage As New Bitmap(aSrcFileName)

        Dim dstImage As WriteableBitmap = RGB_2_Index4(srcImage)

        SaveImage(dstImage, aDstFileName)

    End Sub

    ' Convert input RGB image to Indexed image

    Public Function RGB_2_Index4(aSrcImage As Bitmap) As WriteableBitmap

        ' Calculate target stride

        Dim dstStride As Integer ' Stride of target image
        dstStride = aSrcImage.Width * 2 ' bits
        dstStride = CInt(Math.Ceiling(CDbl(dstStride) / 8.0)) ' bytes

        ' Allocate and get source image data

        Dim srcData As BitmapData = aSrcImage.LockBits(New Rectangle(0, 0, aSrcImage.Width, aSrcImage.Height), ImageLockMode.[ReadOnly], System.Drawing.Imaging.PixelFormat.Format24bppRgb)

        Dim srcPixels As Byte() = New Byte(srcData.Stride * aSrcImage.Height - 1) {} ' Souce image data
        System.Runtime.InteropServices.Marshal.Copy(srcData.Scan0, srcPixels, 0, srcPixels.Length)

        aSrcImage.UnlockBits(srcData)

        ' Allocate target image data

        Dim dstPixels As Byte() = New Byte(dstStride * aSrcImage.Height - 1) {} ' Target image data

        '
        ' Quantization
        '

        Dim dstIndexes As Integer() = New Integer(3) {}

        ' scanline (scan each image row)

        For height_i As Integer = 0 To aSrcImage.Height - 1

            Dim srcLineIndex As Integer = height_i * srcData.Stride
            Dim dstLineIndex As Integer = height_i * dstStride
            Dim dstBlockIndex As Integer = dstLineIndex

            ' Scan block of 4 pixels at a time. 1 pixel has 3 bytes. 3*4=12.
            ' Reason: byte is minimal operating unit. byte has 8 bits. Target image encodes each pixel as 2 bits. 8/2=4.

            Dim i As Integer = 0
            While i < srcData.Stride
                ' Scan 4 pixels

                For j As Integer = 0 To 3
                    If i + j * 3 + 2 < srcData.Stride Then
                        Dim srcBlockIndex As Integer = srcLineIndex + i + j * 3
                        ' BGR ???
                        dstIndexes(j) = GetPaletteIndex(System.Drawing.Color.FromArgb(srcPixels(srcBlockIndex + 2), srcPixels(srcBlockIndex + 1), srcPixels(srcBlockIndex + 0)), mPalette)
                    Else
                        Exit For
                    End If
                Next

                ' Write block to target image
                ' Move pointer to next block

                dstPixels(dstBlockIndex) = CByte(dstIndexes(0) << 6 Or dstIndexes(1) << 4 Or dstIndexes(2) << 2 Or dstIndexes(3))
                dstBlockIndex += 1

                i += 3 * 4
            End While

        Next

        '
        ' Create target image
        '

        Dim dstImage As New WriteableBitmap(aSrcImage.Width, aSrcImage.Height, mDpiX, mDpiY, mTargetPixelFormat, GenerateTargetPalette())

        dstImage.WritePixels(New Int32Rect(0, 0, dstImage.PixelWidth, dstImage.PixelHeight), dstPixels, dstStride, 0)

        Return dstImage

    End Function

    '
    ' Static members
    '

    ReadOnly mPalette As New List(Of System.Drawing.Color)() From { _
        System.Drawing.Color.Black, _
        System.Drawing.Color.White, _
        System.Drawing.Color.Red}

    ' Target(destination) image options

    ReadOnly mDpiX As Integer = 96
    ReadOnly mDpiY As Integer = 96

    ReadOnly mTargetPixelFormat As System.Windows.Media.PixelFormat = System.Windows.Media.PixelFormats.Indexed2 ' cannot be changed

    '
    ' Service functions
    '

    Function GenerateTargetPalette() As BitmapPalette
        Dim dstPalette As New List(Of System.Windows.Media.Color)(mPalette.Count)

        For Each c As System.Drawing.Color In mPalette
            dstPalette.Add(System.Windows.Media.Color.FromArgb(c.A, c.R, c.G, c.B))
        Next

        Return New BitmapPalette(dstPalette)
    End Function

    Sub SaveImage(aImage As WriteableBitmap, aFileName As String)
        Using stream As New FileStream(aFileName, FileMode.Create)
            Dim encoder As BitmapEncoder = New GifBitmapEncoder()

            encoder.Frames.Add(BitmapFrame.Create(aImage))
            encoder.Save(stream)
        End Using
    End Sub

    Function GetPaletteIndex(color As System.Drawing.Color, palette As IList(Of System.Drawing.Color)) As Integer
        ' initializes the best difference, set it for worst possible, it can only get better

        Dim leastDistance As Long = Long.MaxValue
        Dim result As Integer = 0

        For index As Integer = 0 To palette.Count - 1
            Dim targetColor As System.Drawing.Color = palette(index)
            Dim distance As Long = GetColorEuclideanDistance(color, targetColor)

            ' if a difference is zero, we're good because it won't get better
            If distance = 0 Then
                result = index
                Exit For
            End If

            ' if a difference is the best so far, stores it as our best candidate
            If distance < leastDistance Then
                leastDistance = distance
                result = index
            End If
        Next

        Return result
    End Function

    Function GetColorEuclideanDistance(requestedColor As System.Drawing.Color, realColor As System.Drawing.Color) As Long
        Dim componentA As Integer = Convert.ToInt64(requestedColor.R) - realColor.R,
            componentB As Integer = Convert.ToInt64(requestedColor.G) - realColor.G,
            componentC As Integer = Convert.ToInt64(requestedColor.B) - realColor.B

        Return CLng(componentA * componentA + componentB * componentB + componentC * componentC)
    End Function

End Module
