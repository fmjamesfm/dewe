"""
Dewesoft control module.

To use, import the module and call initialise_dewe from the *main thread*.
Then an instance of DewesoftWrapper can be instantiated, and used to access a number of features.

TODO: include option for measurement with storing, and load stored data
"""
import sys
import numpy as np
import pythoncom
from threading import Event, main_thread, current_thread
from win32com.client import Dispatch
from pythoncom import com_error, IID_IDispatch, CoMarshalInterThreadInterfaceInStream, \
    CoGetInterfaceAndReleaseStream

import os
import pint
import time
import atexit
import logging
import pandas as pd

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# Unit registry
ureg = pint.UnitRegistry()
ureg.load_definitions(str(os.path.join(os.path.dirname(__file__), 'mod_defs_en.txt')))

### Constants

FS = 100.0

BLOCK_SIZE = 1000
OVERLAP = 0

CTRL_DO_TRIG = 'Ctrl DO Trig'
MEASURE_TRY_LIMIT = 10

ATYPE_CTLAST = 0
ATYPE_CTOVERLAP = 1
ATYPE_CTTRIGGER = 2
ATYPE_CTNEW = 3

measure_running = Event()

global dw_main


def on_exit():
    global dw_main
    dw_main = None
    pythoncom.CoUninitialize()


atexit.register(on_exit)


def reconnect_dewe():
    logger.info('Creating Dewesoft COM connection...')
    global dw_main
    # global dw_id

    if current_thread() != main_thread():
        raise Exception('Restarting Dewesoft connection must be done in main thread!')
        
    dw_main = Dispatch("Dewesoft.App")
    # dw_id = CoMarshalInterThreadInterfaceInStream(IID_IDispatch, dw_main)
    # dw_id = dw_main.GetCreatedThreadId()
    logger.info('Dewesoft COM connection established.')
    return dw_main


initialise_dewe = reconnect_dewe


def dw_from_id(dw_id):
    pythoncom.CoInitialize()
    return Dispatch(CoGetInterfaceAndReleaseStream(dw_id, IID_IDispatch))


def get_new_dw_id():
    global dw_main
    return CoMarshalInterThreadInterfaceInStream(IID_IDispatch, dw_main)


def DataObj(data, fs, name='', unit=None, **kwargs):
    data = pd.Series(data)
    data.fs = fs
    data.name = name
    data.unit = unit
    for key, val in kwargs.items():
        setattr(data, key, val)

    return data


class DewesoftWrapper(object):
    """
    Wrapper for Dewesoft COM communication.
    """

    def __init__(self, fs=FS, visible=True, notify_window=None, dw_id=None):

        logger.info('Active COM connections: {}'.format(pythoncom._GetInterfaceCount()))

        self.measureThread = None
        self.isMeasuring = False
        self.stopRequest = False
        self.used_channels = {}
        self.channels = []
        self.sampleRate = fs
        self.visible = visible
        self._notify_window = notify_window
        self.used_channels_list = []

        global dw_main

        if current_thread() == main_thread():
            try:
                self.dw = dw_main
            except NameError:
                initialise_dewe()
                self.dw = dw_main
        else:
            try:
                self.dw = dw_from_id(dw_id)
            except NameError:
                raise Exception(
                    'You must first initialise the COM connection in the main thread using initialise_dewe()')

    def reconnect(self):
        """
        Restart COM client connection to Dewesoft (only works in main thread!)
        """
        self.dw = reconnect_dewe()

    def start_dewesoft(self):
        """
        Launch Dewesoft
        """
        sys.stdout.flush()
        logger.info("Launching Dewesoft X...")
        try:
            self.dw.Init()
        except BaseException:
            logger.exception('Failed to initialise Dewesoft')
            sys.stdout.flush()
            return

        sys.stdout.flush()
        logger.info("Dewesoft initialised, applying settings...")
        if self.visible:
            self.dw.Visible = 1
        else:
            self.dw.Visible = 0
        sys.stdout.flush()
        self.dw.Enabled = 1
        sys.stdout.flush()

        self.dw.SuppressMessages = False
        sys.stdout.flush()

        if self.dw.Acquiring:
            logger.info("Dewesoft acquiring, stopping...")
            self.dw.Stop()
            sys.stdout.flush()

        logger.info("Dewesoft setup complete")

    def load_setup(self, path):
        """
        Load Dewesoft setup file.
        """
        logger.info('Loading Dewesoft X setup file: {}'.format(os.path.abspath(path)))
        self.dw.LoadSetup(os.path.abspath(path))
        sys.stdout.flush()
        self.get_input_channels(self.dw)

    def load_stored_data(self):
        """
        Load stored file, export to txt and return results as DataFrame
        """

        self.load_last_file()
        result_path = os.path.abspath('.\\temp_result.txt')
        if os.path.exists(result_path):
            os.remove(result_path)

        self.dw.ExportData(7, 0, result_path)
        self.dw.Measure()

        return pd.read_csv(result_path, sep='\t', header=[11], index_col=0)

    def measure_stop(self):
        """
        Stop measurement and reset measure_running event.
        """
        self.dw.Stop()
        sys.stdout.flush()
        measure_running.clear()

    def measure_timed(self, t):
        """
        Run timed measurement.
        """
        self.measure_free()
        time.sleep(t)
        self.measure_stop()

    def stop(self):
        """
        Stop measurement if running.
        """
        sys.stdout.flush()
        logger.info("Dewesoft stopping...")

        if self.dw.Acquiring:
            logger.info("Dewesoft is in acquisition mode, stopping...")

            try:
                self.dw.Stop()
            except:
                logger.warning("Could not stop Dewesoft!")
            measure_running.clear()
        sys.stdout.flush()

    def restart(self):
        del self.dw
        sys.stdout.flush()
        self.start_dewesoft()

    def reset_channels(self):
        [self.used_channels[s].start() for s in self.used_channels]

    def start_storing(self):
        """
        Start measurement with storing to catemp.dxd file in default Dewesoft dir.
        """
        datapath = self.dw.GetSpecDir(4) + 'catemp.dxd'
        if os.path.exists(datapath):
            os.remove(datapath)
        self.dw.StartStoring(datapath)

    def pause_storing(self, path):
        self.dw.PauseStoring()

    def load_last_file(self):
        self.dw.SaveSetup('')
        self.dw.LoadFile(self.dw.UsedDatafile)

    def measure_free(self):
        """
        Start measurement
        """
        sys.stdout.flush()
        val = self.dw.Start()
        sys.stdout.flush()
        if not val:
            raise Exception('Could not start measurement')
        measure_running.set()

    def stop_measure(self):
        self.dw.Stop()
        sys.stdout.flush()
        measure_running.clear()

    def get_sample_rate(self):
        return self.dw.MeasureSampleRateEx

    def set_sample_rate(self, fs):
        """
        Set measurement sample rate in Hz.
        """
        sys.stdout.flush()
        self.dw.MeasureSampleRateEx = fs
        sys.stdout.flush()
        self.sampleRate = fs
        logger.info(
            "Dewesoft acquisition sample rate set to {0} Hz".format(
                self.dw.MeasureSampleRateEx))
        sys.stdout.flush()

    def get_input_channels(self, dewe):
        self.channels = [
            dewe.Data.AllChannels.Item(i) for i in range(
                dewe.Data.AllChannels.Count)]
        logger.debug("Number of channels: %d" % dewe.Data.AllChannels.Count)
        # used channels
        self.used_channels = {}
        for i in range(0, dewe.Data.UsedChannels.Count):
            ch = dewe.Data.UsedChannels.Item(i)
            sys.stdout.flush()
            name = ch.Name.lower()
            # logger.info("Channel name: {0}".format(name))
            self.used_channels[name] = Channel(
                ch, i, block_size=int(self.sampleRate))
            logger.debug('{} connected'.format(name))
        self.used_channels_list = [self.used_channels[ch]
                                   for ch in self.used_channels]
        return self.used_channels

    def get_data_by_name(self, channel_name):
        for key, ch in list(self.get_input_channels(self.dw).items()):
            if channel_name.lower() in key:
                return self.get_data({key: ch})[0]

    def get_data_by_names(self, channel_names):

        out = []
        for nm in channel_names:
            out.append(self.get_data_by_name(nm))
        return out

    def get_channel_by_name(self, channel_name):
        """
        Get data of channel with name channel_name (case insensitive).
        """
        for key, ch in list(self.get_input_channels(self.dw).items()):
            if channel_name.lower() in key:
                return ch

    def get_data(self, channels, start=None, stop=None):
        """
        Returns data stored in channel buffers as pandas.Series
        """
        logger.debug("Acquiring data Start: {} Stop: {}".format(start, stop))
        self.data = []
        for name in channels:
            if start is None:
                start = 0
            if stop is None:
                stop = -1
            rawdata = channels[name].get_data(start=start, stop=stop)
            logger.debug("Data size: {}".format(len(rawdata)))
            if name == CTRL_DO_TRIG:
                continue
            unit = channels[name].unit
            if not np.any(rawdata) == None:
                self.data.append(
                    DataObj(
                        rawdata,
                        self.sampleRate,
                        name=name,
                        unit=unit))
        self.channels_with_data = [
            channels[s] for s in channels if not np.any(
                channels[s].data) == None]
        return self.data

    def __del__(self):
        if current_thread() != main_thread():
            pythoncom.CoUninitialize()

        del self


def convert_dewe_unit(unit):
    """
    Convert unit string to pint unit object from ureg.
    """
    splitted = unit.split('2')
    newunit = splitted[0]
    for i in range(1, len(splitted)):
        newunit += '^2'
        newunit += splitted[i]
    if newunit == '\xb0C':
        newunit = 'degC'
    try:
        output = ureg(newunit).units
    except BaseException:
        logger.error('Unit conversion error')
        output = newunit
    return output


class Channel(object):
    """
    Contains Dewesoft DCOM channel and connection. Used to acquire data from a channel.
    """

    def __init__(self, channel, index, atype=ATYPE_CTNEW,
                 block_size=BLOCK_SIZE, overlap=0):
        self.channel = channel
        self.connection = channel.CreateConnection()
        self.index = index
        sys.stdout.flush()
        self.scale = self.channel.Scale
        self.offset = self.channel.Offset
        self.name = channel.Name
        self.unit = convert_dewe_unit(channel.Unit_)
        self.data = None
        self.AType = atype
        self.blockSize = block_size
        self.overlap = overlap
        self.connection.AType = atype
        self.connection.BlockSize = block_size
        self.connection.Overlap = overlap
        self.numBlocks = self.connection.numBlocks
        self.start()

    def start(self):
        self.connection.Start()
        sys.stdout.flush()

    def get_data(self, start=0, stop=-1):
        maxLength = self.channel.DBBufSize
        n_values = self.connection.NumValues

        if not (start == 0 and stop == -1):
            startInd = start % maxLength
            stopInd = stop % maxLength
        else:
            startInd = start
            stopInd = stop

        self.data = np.array(self.connection.GetDataValues(n_values))

        logger.debug(
            "Channel {} direct buffer size: {}, data length: {}".format(
                self.name, maxLength, len(
                    self.data)))
        logger.debug(
            'Start index: {} Stop index {}, N. values: {}'.format(
                startInd, stopInd, n_values))
        try:
            self.data[0]
        except BaseException:
            logger.warning(
                'No data retrieved from channel {0}'.format(
                    self.name))
            return
        self.start()
        if startInd >= stopInd and stopInd > 0:
            output = np.concatenate(
                (self.data[startInd:], self.data[0:stopInd]))
        else:
            output = self.data[startInd:stopInd]

        return output


def _test():
    initialise_dewe()
    dw = DewesoftWrapper()
    # dw.load_setup('test.dxs')

    dw.start_storing()
    time.sleep(3)

    dw.measure_stop()
    dt = dw.load_stored_data()
    print(dt.columns)


if __name__ == '__main__':
    lgr = logging.getLogger()
    lgr.setLevel(logging.INFO)
    logging.basicConfig()
    _test()
