Basic micropython library to control the OLED SSD1306 128x64 spi with a micro:bit
#################################################################################

This library allows the micro:bit to control the typical low cost 0,96" OLED display sold in Amazon and eBay connected to the default spi pins of the micro:bit. Some sort of breakout is required.

You should connect D0 to 13, D1 to 15, RES to 14 and DC to 16. You also must connect the device’s ground to the micro:bit ground (pin GND). 

This library uses the full resolution of the OLED, due to some optimizations that can be done when using SPI instead of I2C.


   .. image:: ./images/ssd1306spi_sm.jpg
      :width: 100%
      :align: center
      
.. contents::

.. section-numbering::


Main features
=============

* Load a 128x64 bitmap file
* Set and get pixel value
* Sprites
* Text
* Sample programs demonstrating the different functions


Preparation and displaying of a bitmap image
============================================

1. Create a bitmap with an image editor with only 2 bits per pixel (black and white) 
2. Use the LCDAssistant (http://en.radzio.dxp.pl/bitmap_converter/) to generate the hex data. 
3. Copy the hex data into the bitmap_converter.py file and run it on a computer.
4. Flash a completely empty file from mu.
5. Copy the generated file to the micro:bit using the file transfer function in mu
6. Create a main.py file, import sdd1306spi and use the function show_bitmap to display the file
7. Move the files main.py, sdd1306.py and sdd1306_bitmap.py to the micro:bit with the file transfer function in mu
8. Reset the micro:bit or press CTRL+D in the Repl.


Library usage
=============


initilization
+++++++++++++++++++++++


You have to instantiate the SSD1306 object before using the display. This puts the display in its reset status.

.. code-block:: python

   from ssd1306spi import SSD1306
   
   oled = SSD1306()


clear_oled()
+++++++++++++++++++++++


You will typically use this function after instantiating the object, in order to make sure that the display is blank at the beginning. 


show_bitmap(filename)
+++++++++++++++++++++++


Displays on the OLED screen the image stored in the file *filename*. The image has to be encode as described in the previous section.

.. code-block:: python

   from ssd1306spi import SSD1306
   
   oled = SSD1306()
   oled.clear_oled()
   oled.show_bitmap("microbit_logo")

set_px(x, y, color, draw=1)
+++++++++++++++++++++++++++++


Paints the pixel at position x, y (of a 64x32 coordinate system) with the corresponding color (0 dark or 1 lighted). 
If the optional parameter **draw** is set to 0 the screen will not be refreshed and **draw_screen()** needs to be called at a later stage, since multiple screen refreshes can be time consuming. This allows setting different pixels in the buffer without refreshing the screen, and finally refresh the display with the content of the buffer.

.. code-block:: python

   from ssd1306spi import SSD1306
   
   oled = SSD1306()
   oled.clear_oled()
   oled.set_px(10,10,1)
   oled.set_px(20,20,0,0)
   oled.draw_screen()


get_px(x, y)
++++++++++++


Returns the color of the given pixel (0 dark 1 lighted)

.. code-block:: python

   from ssd1306spi import SSD1306
   
   oled = SSD1306()
   oled.clear_oled()
   color = oled.get_px(10,10)


draw_sprite(x, y, stamp, color, draw=1)
++++++++++++++++++++++++++++++++++++++

Draws the sprite on the screen at the pixel position x, y. The sprite will be printed using **OR** if color is 1 and **AND NOT** if color is 0, effectively removing the sprite when color=0.

.. code-block:: python

   from ssd1306spi import SSD1306
   
   oled = SSD1306()
   oled.clear_oled()
   sprt = b'\xAE\xA4\xD5\xF0\xA8\x3F\xD3\x00\x00\x8D'
   oled.draw_sprite(0, 0, sprt, 1, 0)
   

When drawing a sprite, the contents of the screen just before the first column of the stamp and the content of the screen just after the last column of the sprite is also redrawn. This is done to allow using a function like this to perform a simple movement of a sprite:

.. code-block:: python

    def move_sprite(oled, x1, y1, x2, y2, sprt):
      oled.draw_sprite(x1, y1, sprt, 0, 0)
      oled.draw_sprite(x2, y2, sprt, 1, 1)
      
      
The previous function removes a sprite at position x1,y1 and redraws it at position x2, y2. Note that the first draw_sprite() does not refresh the screen. The screen is only refreshed once, with the second draw_sprte(). If the sprite is 5x5 and it is centered within the 8x7 area, the sprite will be properly updated if the distance between the two coordinates is maximum one pixel.

