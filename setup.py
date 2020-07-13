from setuptools import setup
setup(
    name='address_geocoder',
    version='0.0.1',
    entry_points={
        'console_scripts': [
            'geocoder=address_geocoder:run'
        ]
    }
)